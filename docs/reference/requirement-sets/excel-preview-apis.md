---
title: APIs de visualização do JavaScript para Excel
description: Detalhes sobre as próximas Excel APIs JavaScript.
ms.date: 09/16/2021
ms.prod: excel
ms.localizationpriority: medium
ms.openlocfilehash: bd36d9ba1be4e9e0caafdd49e63d8e7cdea01c59
ms.sourcegitcommit: a854a2fd2ad9f379a3ef712f307e0b1bb9b5b00d
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 09/22/2021
ms.locfileid: "59474347"
---
# <a name="excel-javascript-preview-apis"></a>APIs de visualização do JavaScript para Excel

As novas APIs do JavaScript para Excel são introduzidas pela primeira vez na "visualização" e, posteriormente, tornam-se parte de um conjunto específico de requisitos numerados, após passarem por vários testes e após a recolha das opiniões de usuários.

[!INCLUDE [Information about using preview APIs](../../includes/using-preview-apis-host.md)]

A tabela a seguir fornece um resumo conciso das APIs, enquanto a tabela de lista [de API](#api-list) subsequente fornece uma lista detalhada.

| Área de recurso | Descrição | Objetos relevantes |
|:--- |:--- |:--- |
| Tabelas de dados de gráfico | Controlar a aparência, a formatação e a visibilidade das tabelas de dados nos gráficos. | [Chart,](/javascript/api/excel/excel.chart) [ChartDataTable,](/javascript/api/excel/excel.chartdatatable) [ChartDataTableFormat](/javascript/api/excel/excel.chartdatatableformat) |
| Tipos de dados personalizados | Uma extensão de tipos de dados Excel existentes, incluindo suporte para números formatados e imagens da Web. | [BooleanCellValue](/javascript/api/excel/excel.booleancellvalue), [CellValueAttributionAttributes](/javascript/api/excel/excel.cellvalueattributionattributes), [CellValueProviderAttributes](/javascript/api/excel/excel.cellvalueproviderattributes), [DoubleCellValue](/javascript/api/excel/excel.doublecellvalue), [EmptyCellValue](/javascript/api/excel/excel.emptycellvalue), [FormattedNumberCellValue](/javascript/api/excel/excel.formattednumbercellvalue), [StringCellValue](/javascript/api/excel/excel.stringcellvalue), [ValueTypeNotAvailableCellValue](/javascript/api/excel/excel.valuetypenotavailablecellvalue), [WebImageCellValue](/javascript/api/excel/excel.webimagecellvalue) |
| Erros de tipos de dados personalizados| Objetos de erro que suportam tipos de dados personalizados. | [BlockedErrorCellValue](/javascript/api/excel/excel.blockederrorcellvalue), [BusyErrorCellValue](/javascript/api/excel/excel.busyerrorcellvalue), [CalcErrorCellValue](/javascript/api/excel/excel.calcerrorcellvalue), [ConnectErrorCellValue](/javascript/api/excel/excel.connecterrorcellvalue), [Div0ErrorCellValue](/javascript/api/excel/excel.div0errorcellvalue), [FieldErrorCellValue](/javascript/api/excel/excel.fielderrorcellvalue), [GettingDataErrorCellValue](/javascript/api/excel/excel.gettingdataerrorcellvalue), [NaErrorCellValue](/javascript/api/excel/excel.naerrorcellvalue), [NameErrorCellValue](/javascript/api/excel/excel.nameerrorcellvalue), [NullErrorCellValue](/javascript/api/excel/excel.nullerrorcellvalue), [NumErrorCellValue](/javascript/api/excel/excel.numerrorcellvalue), [RefErrorCellValue](/javascript/api/excel/excel.referrorcellvalue), [SpillErrorCellValue](/javascript/api/excel/excel.spillerrorcellvalue), [ValueErrorCellValue](/javascript/api/excel/excel.valueerrorcellvalue)|
| Tarefas do documento | Transforme os comentários em tarefas atribuídas aos usuários. | [DocumentTask](/javascript/api/excel/excel.documenttask) |
| Identidades | Gerencie identidades de usuário, incluindo nome de exibição e endereço de email. | [Identity](/javascript/api/excel/excel.identity), [IdentityCollection](/javascript/api/excel/excel.identitycollection), [IdentityEntity](/javascript/api/excel/excel.identityentity) |
| Tipos de dados vinculados | Adiciona suporte para tipos de dados conectados Excel de fontes externas. | [LinkedDataType](/javascript/api/excel/excel.linkeddatatype)|
| Table styles | Fornece controle para fonte, borda, cor de preenchimento e outros aspectos dos estilos de tabela. | [Tabela,](/javascript/api/excel/excel.table) [Tabela Dinâmica,](/javascript/api/excel/excel.pivottable) [Slicer](/javascript/api/excel/excel.slicer) |
| Consultas | Recupere atributos de consulta, como nome, data de atualização e contagem de consultas. | [Consulta](/javascript/api/excel/excel.query), [QueryCollection](/javascript/api/excel/excel.querycollection)|
| Proteção de planilha | Impedir que usuários não autorizados mudem para intervalos especificados em uma planilha. | [WorksheetProtection](/javascript/api/excel/excel.worksheetprotection), [WorksheetProtectionChangedEventArgs](/javascript/api/excel/excel.worksheetprotectionchangedeventargs), [AllowEditRange](/javascript/api/excel/excel.alloweditrange), [AllowEditRangeCollection](/javascript/api/excel/excel.alloweditrangecollection), [AllowEditRangeOptions](/javascript/api/excel/excel.alloweditrangeoptions) |

## <a name="api-list"></a>Lista de API

A tabela a seguir lista as Excel APIs JavaScript atualmente em visualização. Para uma lista completa de todas as EXCEL JavaScript (incluindo APIs de visualização e APIs lançadas anteriormente), consulte todas as [APIs JavaScript](/javascript/api/excel?view=excel-js-preview&preserve-view=true)Excel JavaScript .

| Classe | Campos | Descrição |
|:---|:---|:---|
|[AllowEditRange](/javascript/api/excel/excel.alloweditrange)|[address](/javascript/api/excel/excel.alloweditrange#address)|Especifica o intervalo associado ao objeto.|
||[delete()](/javascript/api/excel/excel.alloweditrange#delete__)|Exclui esse objeto do `AllowEditRangeCollection` .|
||[pauseProtection(password?: string)](/javascript/api/excel/excel.alloweditrange#pauseProtection_password_)|Pausa a proteção da planilha para o `AllowEditRange` objeto dado para o usuário em uma determinada sessão.|
||[isPasswordProtected](/javascript/api/excel/excel.alloweditrange#isPasswordProtected)|Especifica se a `AllowEditRange` senha está protegida.|
||[setPassword(password?: string)](/javascript/api/excel/excel.alloweditrange#setPassword_password_)|Altera a senha associada ao `AllowEditRange` .|
||[title](/javascript/api/excel/excel.alloweditrange#title)|Especifica o título do objeto.|
|[AllowEditRangeCollection](/javascript/api/excel/excel.alloweditrangecollection)|[add(title: string, rangeAddress: string, options?: Excel. AllowEditRangeOptions)](/javascript/api/excel/excel.alloweditrangecollection#add_title__rangeAddress__options_)|Adiciona um `AllowEditRange` objeto à coleção.|
||[getCount()](/javascript/api/excel/excel.alloweditrangecollection#getCount__)|Retorna o número `AllowEditRange` de objetos na coleção.|
||[getItem(key: string)](/javascript/api/excel/excel.alloweditrangecollection#getItem_key_)|Obtém `AllowEditRange` o objeto pelo título.|
||[getItemAt(index: number)](/javascript/api/excel/excel.alloweditrangecollection#getItemAt_index_)|Retorna um `AllowEditRange` objeto pelo índice na coleção.|
||[getItemOrNullObject(key: string)](/javascript/api/excel/excel.alloweditrangecollection#getItemOrNullObject_key_)|Obtém `AllowEditRange` o objeto pelo título.|
||[pauseProtection(password: string)](/javascript/api/excel/excel.alloweditrangecollection#pauseProtection_password_)|Pausa a proteção da planilha para todos os objetos da coleção que têm a `AllowEditRange` senha dada para o usuário em uma determinada sessão.|
||[items](/javascript/api/excel/excel.alloweditrangecollection#items)|Obtém os itens filhos carregados nesta coleção.|
|[AllowEditRangeOptions](/javascript/api/excel/excel.alloweditrangeoptions)|[password](/javascript/api/excel/excel.alloweditrangeoptions#password)|A senha associada ao `AllowEditRange` .|
|[BlockedErrorCellValue](/javascript/api/excel/excel.blockederrorcellvalue)|[errorSubType](/javascript/api/excel/excel.blockederrorcellvalue#errorSubType)|Representa o tipo de `BlockedErrorCellValue` .|
||[errorType](/javascript/api/excel/excel.blockederrorcellvalue#errorType)|Representa o tipo de `ErrorCellValue` .|
||[primitive](/javascript/api/excel/excel.blockederrorcellvalue#primitive)|Representa o valor que seria retornado por `Range.values` uma célula com esse valor.|
||[primitiveType](/javascript/api/excel/excel.blockederrorcellvalue#primitiveType)|Representa o valor que seria retornado por `Range.valueTypes` uma célula com esse valor.|
||[type](/javascript/api/excel/excel.blockederrorcellvalue#type)|Representa o tipo desse valor de célula.|
|[BooleanCellValue](/javascript/api/excel/excel.booleancellvalue)|[primitive](/javascript/api/excel/excel.booleancellvalue#primitive)|Representa o valor que seria retornado por `Range.values` uma célula com esse valor.|
||[primitiveType](/javascript/api/excel/excel.booleancellvalue#primitiveType)|Representa o valor que seria retornado por `Range.valueTypes` uma célula com esse valor.|
||[type](/javascript/api/excel/excel.booleancellvalue#type)|Representa o tipo desse valor de célula.|
|[BusyErrorCellValue](/javascript/api/excel/excel.busyerrorcellvalue)|[errorSubType](/javascript/api/excel/excel.busyerrorcellvalue#errorSubType)|Representa o tipo de `BusyErrorCellValue` .|
||[errorType](/javascript/api/excel/excel.busyerrorcellvalue#errorType)|Representa o tipo de `ErrorCellValue` .|
||[primitive](/javascript/api/excel/excel.busyerrorcellvalue#primitive)|Representa o valor que seria retornado por `Range.values` uma célula com esse valor.|
||[primitiveType](/javascript/api/excel/excel.busyerrorcellvalue#primitiveType)|Representa o valor que seria retornado por `Range.valueTypes` uma célula com esse valor.|
||[type](/javascript/api/excel/excel.busyerrorcellvalue#type)|Representa o tipo desse valor de célula.|
|[CalcErrorCellValue](/javascript/api/excel/excel.calcerrorcellvalue)|[errorSubType](/javascript/api/excel/excel.calcerrorcellvalue#errorSubType)|Representa o tipo de `CalcErrorCellValue` .|
||[errorType](/javascript/api/excel/excel.calcerrorcellvalue#errorType)|Representa o tipo de `ErrorCellValue` .|
||[primitive](/javascript/api/excel/excel.calcerrorcellvalue#primitive)|Representa o valor que seria retornado por `Range.values` uma célula com esse valor.|
||[primitiveType](/javascript/api/excel/excel.calcerrorcellvalue#primitiveType)|Representa o valor que seria retornado por `Range.valueTypes` uma célula com esse valor.|
||[type](/javascript/api/excel/excel.calcerrorcellvalue#type)|Representa o tipo desse valor de célula.|
|[CellValueAttributionAttributes](/javascript/api/excel/excel.cellvalueattributionattributes)|[licenseAddress](/javascript/api/excel/excel.cellvalueattributionattributes#licenseAddress)|Representa uma URL para uma licença ou fonte que descreve como essa propriedade pode ser usada.|
||[licenseText](/javascript/api/excel/excel.cellvalueattributionattributes#licenseText)|Representa um nome para a licença que rege essa propriedade.|
||[sourceAddress](/javascript/api/excel/excel.cellvalueattributionattributes#sourceAddress)|Representa uma URL para a origem do `CellValue` .|
||[sourceText](/javascript/api/excel/excel.cellvalueattributionattributes#sourceText)|Representa um nome para a origem do `CellValue` .|
|[CellValueProviderAttributes](/javascript/api/excel/excel.cellvalueproviderattributes)|[description](/javascript/api/excel/excel.cellvalueproviderattributes#description)|Representa a propriedade de descrição do provedor usada no exibição de cartão se nenhum logotipo for especificado.|
||[logoSourceAddress](/javascript/api/excel/excel.cellvalueproviderattributes#logoSourceAddress)|Representa uma URL usada para baixar uma imagem que será usada como um logotipo no exibição de cartão.|
||[logoTargetAddress](/javascript/api/excel/excel.cellvalueproviderattributes#logoTargetAddress)|Representa uma URL que será o destino de navegação se o usuário clicar no elemento logo no modo de exibição de cartão.|
|[ChangeDirectionState](/javascript/api/excel/excel.changedirectionstate)|[deleteShiftDirection](/javascript/api/excel/excel.changedirectionstate#deleteShiftDirection)|Representa a direção (como para cima ou para a esquerda) que as células restantes serão deslocadas quando uma célula ou células são excluídas.|
||[insertShiftDirection](/javascript/api/excel/excel.changedirectionstate#insertShiftDirection)|Representa a direção (como para baixo ou para a direita) que as células existentes mudarão quando uma nova célula ou células são inseridas.|
|[Chart](/javascript/api/excel/excel.chart)|[getDataTable()](/javascript/api/excel/excel.chart#getDataTable__)|Obtém a tabela de dados no gráfico.|
||[getDataTableOrNullObject()](/javascript/api/excel/excel.chart#getDataTableOrNullObject__)|Obtém a tabela de dados no gráfico.|
|[ChartDataTable](/javascript/api/excel/excel.chartdatatable)|[format](/javascript/api/excel/excel.chartdatatable#format)|Representa o formato de uma tabela de dados do gráfico, que inclui o formato de preenchimento, fonte e borda.|
||[showHorizontalBorder](/javascript/api/excel/excel.chartdatatable#showHorizontalBorder)|Especifica se a borda horizontal da tabela de dados deve ser exibida.|
||[showLegendKey](/javascript/api/excel/excel.chartdatatable#showLegendKey)|Especifica se a chave de legenda da tabela de dados deve ser apresentada.|
||[showOutlineBorder](/javascript/api/excel/excel.chartdatatable#showOutlineBorder)|Especifica se a borda de contorno da tabela de dados deve ser exibida.|
||[showVerticalBorder](/javascript/api/excel/excel.chartdatatable#showVerticalBorder)|Especifica se a borda vertical da tabela de dados deve ser exibida.|
||[visible](/javascript/api/excel/excel.chartdatatable#visible)|Especifica se é preciso mostrar a tabela de dados do gráfico.|
|[ChartDataTableFormat](/javascript/api/excel/excel.chartdatatableformat)|[borda](/javascript/api/excel/excel.chartdatatableformat#border)|Representa o formato de borda da tabela de dados do gráfico, que inclui cor, estilo de linha e peso.|
||[fill](/javascript/api/excel/excel.chartdatatableformat#fill)|Representa o formato de preenchimento de um objeto, que inclui informações sobre a formatação da tela de fundo.|
||[font](/javascript/api/excel/excel.chartdatatableformat#font)|Representa os atributos de fonte (como nome da fonte, tamanho da fonte e cor) do objeto atual.|
|[Comment](/javascript/api/excel/excel.comment)|[assignTask(assignee: Identity)](/javascript/api/excel/excel.comment#assignTask_assignee_)|Atribui a tarefa anexada ao comentário ao usuário dado como um destinatário.|
||[getTask()](/javascript/api/excel/excel.comment#getTask__)|Obtém a tarefa associada a este comentário.|
||[getTaskOrNullObject()](/javascript/api/excel/excel.comment#getTaskOrNullObject__)|Obtém a tarefa associada a este comentário.|
|[CommentCollection](/javascript/api/excel/excel.commentcollection)|[getItemOrNullObject(commentId: string)](/javascript/api/excel/excel.commentcollection#getItemOrNullObject_commentId_)|Obtém um comentário da coleção com base em seu ID.|
|[CommentReply](/javascript/api/excel/excel.commentreply)|[assignTask(assignee: Identity)](/javascript/api/excel/excel.commentreply#assignTask_assignee_)|Atribui a tarefa anexada ao comentário ao usuário determinado como o único destinatário.|
||[getTask()](/javascript/api/excel/excel.commentreply#getTask__)|Obtém a tarefa associada ao thread desta resposta de comentário.|
||[getTaskOrNullObject()](/javascript/api/excel/excel.commentreply#getTaskOrNullObject__)|Obtém a tarefa associada ao thread desta resposta de comentário.|
|[CommentReplyCollection](/javascript/api/excel/excel.commentreplycollection)|[getItemOrNullObject(commentReplyId: string)](/javascript/api/excel/excel.commentreplycollection#getItemOrNullObject_commentReplyId_)|Retorna uma resposta de comentário identificada pela respectiva ID.|
|[ConditionalFormatCollection](/javascript/api/excel/excel.conditionalformatcollection)|[getItemOrNullObject(id: string)](/javascript/api/excel/excel.conditionalformatcollection#getItemOrNullObject_id_)|Retorna um formato condicional identificado por sua ID.|
|[ConnectErrorCellValue](/javascript/api/excel/excel.connecterrorcellvalue)|[errorSubType](/javascript/api/excel/excel.connecterrorcellvalue#errorSubType)|Representa o tipo de `ConnectErrorCellValue` .|
||[errorType](/javascript/api/excel/excel.connecterrorcellvalue#errorType)|Representa o tipo de `ErrorCellValue` .|
||[primitive](/javascript/api/excel/excel.connecterrorcellvalue#primitive)|Representa o valor que seria retornado por `Range.values` uma célula com esse valor.|
||[primitiveType](/javascript/api/excel/excel.connecterrorcellvalue#primitiveType)|Representa o valor que seria retornado por `Range.valueTypes` uma célula com esse valor.|
||[type](/javascript/api/excel/excel.connecterrorcellvalue#type)|Representa o tipo desse valor de célula.|
|[Div0ErrorCellValue](/javascript/api/excel/excel.div0errorcellvalue)|[errorType](/javascript/api/excel/excel.div0errorcellvalue#errorType)|Representa o tipo de `ErrorCellValue` .|
||[primitive](/javascript/api/excel/excel.div0errorcellvalue#primitive)|Representa o valor que seria retornado por `Range.values` uma célula com esse valor.|
||[primitiveType](/javascript/api/excel/excel.div0errorcellvalue#primitiveType)|Representa o valor que seria retornado por `Range.valueTypes` uma célula com esse valor.|
||[type](/javascript/api/excel/excel.div0errorcellvalue#type)|Representa o tipo desse valor de célula.|
|[DocumentTask](/javascript/api/excel/excel.documenttask)|[percentComplete](/javascript/api/excel/excel.documenttask#percentComplete)|Especifica a porcentagem de conclusão da tarefa.|
||[priority](/javascript/api/excel/excel.documenttask#priority)|Especifica a prioridade da tarefa.|
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
||[priority](/javascript/api/excel/excel.documenttaskchange#priority)|Representa a prioridade da tarefa.|
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
|[DoubleCellValue](/javascript/api/excel/excel.doublecellvalue)|[primitive](/javascript/api/excel/excel.doublecellvalue#primitive)|Representa o valor que seria retornado por `Range.values` uma célula com esse valor.|
||[primitiveType](/javascript/api/excel/excel.doublecellvalue#primitiveType)|Representa o valor que seria retornado por `Range.valueTypes` uma célula com esse valor.|
||[type](/javascript/api/excel/excel.doublecellvalue#type)|Representa o tipo desse valor de célula.|
|[EmptyCellValue](/javascript/api/excel/excel.emptycellvalue)|[primitive](/javascript/api/excel/excel.emptycellvalue#primitive)|Representa o valor que seria retornado por `Range.values` uma célula com esse valor.|
||[primitiveType](/javascript/api/excel/excel.emptycellvalue#primitiveType)|Representa o valor que seria retornado por `Range.valueTypes` uma célula com esse valor.|
||[type](/javascript/api/excel/excel.emptycellvalue#type)|Representa o tipo desse valor de célula.|
|[FieldErrorCellValue](/javascript/api/excel/excel.fielderrorcellvalue)|[errorSubType](/javascript/api/excel/excel.fielderrorcellvalue#errorSubType)|Representa o tipo de `FieldErrorCellValue` .|
||[errorType](/javascript/api/excel/excel.fielderrorcellvalue#errorType)|Representa o tipo de `ErrorCellValue` .|
||[primitive](/javascript/api/excel/excel.fielderrorcellvalue#primitive)|Representa o valor que seria retornado por `Range.values` uma célula com esse valor.|
||[primitiveType](/javascript/api/excel/excel.fielderrorcellvalue#primitiveType)|Representa o valor que seria retornado por `Range.valueTypes` uma célula com esse valor.|
||[type](/javascript/api/excel/excel.fielderrorcellvalue#type)|Representa o tipo desse valor de célula.|
|[FormattedNumberCellValue](/javascript/api/excel/excel.formattednumbercellvalue)|[numberFormat](/javascript/api/excel/excel.formattednumbercellvalue#numberFormat)|Retorna a cadeia de caracteres de formato de número usada para exibir esse valor.|
||[primitive](/javascript/api/excel/excel.formattednumbercellvalue#primitive)|Representa o valor que seria retornado por `Range.values` uma célula com esse valor.|
||[primitiveType](/javascript/api/excel/excel.formattednumbercellvalue#primitiveType)|Representa o valor que seria retornado por `Range.valueTypes` uma célula com esse valor.|
||[type](/javascript/api/excel/excel.formattednumbercellvalue#type)|Representa o tipo desse valor de célula.|
|[GettingDataErrorCellValue](/javascript/api/excel/excel.gettingdataerrorcellvalue)|[errorType](/javascript/api/excel/excel.gettingdataerrorcellvalue#errorType)|Representa o tipo de `ErrorCellValue` .|
||[primitive](/javascript/api/excel/excel.gettingdataerrorcellvalue#primitive)|Representa o valor que seria retornado por `Range.values` uma célula com esse valor.|
||[primitiveType](/javascript/api/excel/excel.gettingdataerrorcellvalue#primitiveType)|Representa o valor que seria retornado por `Range.valueTypes` uma célula com esse valor.|
||[type](/javascript/api/excel/excel.gettingdataerrorcellvalue#type)|Representa o tipo desse valor de célula.|
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
|[NaErrorCellValue](/javascript/api/excel/excel.naerrorcellvalue)|[errorType](/javascript/api/excel/excel.naerrorcellvalue#errorType)|Representa o tipo de `ErrorCellValue` .|
||[primitive](/javascript/api/excel/excel.naerrorcellvalue#primitive)|Representa o valor que seria retornado por `Range.values` uma célula com esse valor.|
||[primitiveType](/javascript/api/excel/excel.naerrorcellvalue#primitiveType)|Representa o valor que seria retornado por `Range.valueTypes` uma célula com esse valor.|
||[type](/javascript/api/excel/excel.naerrorcellvalue#type)|Representa o tipo desse valor de célula.|
|[NameErrorCellValue](/javascript/api/excel/excel.nameerrorcellvalue)|[errorType](/javascript/api/excel/excel.nameerrorcellvalue#errorType)|Representa o tipo de `ErrorCellValue` .|
||[primitive](/javascript/api/excel/excel.nameerrorcellvalue#primitive)|Representa o valor que seria retornado por `Range.values` uma célula com esse valor.|
||[primitiveType](/javascript/api/excel/excel.nameerrorcellvalue#primitiveType)|Representa o valor que seria retornado por `Range.valueTypes` uma célula com esse valor.|
||[type](/javascript/api/excel/excel.nameerrorcellvalue#type)|Representa o tipo desse valor de célula.|
|[NamedSheetViewCollection](/javascript/api/excel/excel.namedsheetviewcollection)|[getItemOrNullObject(key: string)](/javascript/api/excel/excel.namedsheetviewcollection#getItemOrNullObject_key_)|Obtém uma exibição de planilha usando seu nome.|
|[NullErrorCellValue](/javascript/api/excel/excel.nullerrorcellvalue)|[errorType](/javascript/api/excel/excel.nullerrorcellvalue#errorType)|Representa o tipo de `ErrorCellValue` .|
||[primitive](/javascript/api/excel/excel.nullerrorcellvalue#primitive)|Representa o valor que seria retornado por `Range.values` uma célula com esse valor.|
||[primitiveType](/javascript/api/excel/excel.nullerrorcellvalue#primitiveType)|Representa o valor que seria retornado por `Range.valueTypes` uma célula com esse valor.|
||[type](/javascript/api/excel/excel.nullerrorcellvalue#type)|Representa o tipo desse valor de célula.|
|[NumErrorCellValue](/javascript/api/excel/excel.numerrorcellvalue)|[errorType](/javascript/api/excel/excel.numerrorcellvalue#errorType)|Representa o tipo de `ErrorCellValue` .|
||[primitive](/javascript/api/excel/excel.numerrorcellvalue#primitive)|Representa o valor que seria retornado por `Range.values` uma célula com esse valor.|
||[primitiveType](/javascript/api/excel/excel.numerrorcellvalue#primitiveType)|Representa o valor que seria retornado por `Range.valueTypes` uma célula com esse valor.|
||[type](/javascript/api/excel/excel.numerrorcellvalue#type)|Representa o tipo desse valor de célula.|
|[PivotLayout](/javascript/api/excel/excel.pivotlayout)|[getCell(dataHierarchy: DataPivotHierarchy \| string, rowItems: Array<PivotItem \| string>, columnItems: Array<PivotItem \| string>)](/javascript/api/excel/excel.pivotlayout#getCell_dataHierarchy__rowItems__columnItems_)|Obtém uma célula exclusiva na tabela dinâmica com base em uma hierarquia de dados, bem como os itens de linha e coluna de suas respectivas hierarquias.|
||[pivotStyle](/javascript/api/excel/excel.pivotlayout#pivotStyle)|O estilo aplicado à Tabela Dinâmica.|
||[setStyle(style: string \| PivotTableStyle \| BuiltInPivotTableStyle)](/javascript/api/excel/excel.pivotlayout#setStyle_style_)|Define o estilo aplicado à Tabela Dinâmica.|
|[PivotTableScopedCollection](/javascript/api/excel/excel.pivottablescopedcollection)|[getFirstOrNullObject()](/javascript/api/excel/excel.pivottablescopedcollection#getFirstOrNullObject__)|Obtém a primeira Tabela Dinâmica da coleção.|
|[Query](/javascript/api/excel/excel.query)|[error](/javascript/api/excel/excel.query#error)|Obtém a mensagem de erro de consulta de quando a consulta foi atualizada pela última vez.|
||[loadedTo](/javascript/api/excel/excel.query#loadedTo)|Obtém a consulta carregada para o tipo de objeto.|
||[loadedToDataModel](/javascript/api/excel/excel.query#loadedToDataModel)|Especifica se a consulta foi carregada para o modelo de dados.|
||[name](/javascript/api/excel/excel.query#name)|Obtém o nome da consulta.|
||[refreshDate](/javascript/api/excel/excel.query#refreshDate)|Obtém a data e a hora em que a consulta foi atualizada pela última vez.|
||[rowsLoadedCount](/javascript/api/excel/excel.query#rowsLoadedCount)|Obtém o número de linhas que foram carregadas quando a consulta foi atualizada pela última vez.|
|[QueryCollection](/javascript/api/excel/excel.querycollection)|[getCount()](/javascript/api/excel/excel.querycollection#getCount__)|Obtém o número de consultas na guia de trabalho.|
||[getItem(key: string)](/javascript/api/excel/excel.querycollection#getItem_key_)|Obtém uma consulta da coleção com base em seu nome.|
||[items](/javascript/api/excel/excel.querycollection#items)|Obtém os itens filhos carregados nesta coleção.|
|[Range](/javascript/api/excel/excel.range)|[getDependents()](/javascript/api/excel/excel.range#getDependents__)|Retorna um objeto que representa o intervalo que contém todos os dependentes de uma célula na mesma planilha ou `WorkbookRangeAreas` em várias planilhas.|
||[getPrecedents()](/javascript/api/excel/excel.range#getPrecedents__)|Retorna um objeto que representa o intervalo que contém todos os precedentes de uma célula na mesma planilha ou `WorkbookRangeAreas` em várias planilhas.|
|[RefErrorCellValue](/javascript/api/excel/excel.referrorcellvalue)|[errorSubType](/javascript/api/excel/excel.referrorcellvalue#errorSubType)|Representa o tipo de `RefErrorCellValue` .|
||[errorType](/javascript/api/excel/excel.referrorcellvalue#errorType)|Representa o tipo de `ErrorCellValue` .|
||[primitive](/javascript/api/excel/excel.referrorcellvalue#primitive)|Representa o valor que seria retornado por `Range.values` uma célula com esse valor.|
||[primitiveType](/javascript/api/excel/excel.referrorcellvalue#primitiveType)|Representa o valor que seria retornado por `Range.valueTypes` uma célula com esse valor.|
||[type](/javascript/api/excel/excel.referrorcellvalue#type)|Representa o tipo desse valor de célula.|
|[RefreshModeChangedEventArgs](/javascript/api/excel/excel.refreshmodechangedeventargs)|[refreshMode](/javascript/api/excel/excel.refreshmodechangedeventargs#refreshMode)|O modo de atualização do tipo de dados vinculado.|
||[serviceId](/javascript/api/excel/excel.refreshmodechangedeventargs#serviceId)|A ID exclusiva do objeto cujo modo de atualização foi alterado.|
||[source](/javascript/api/excel/excel.refreshmodechangedeventargs#source)|Obtém a origem do evento.|
||[tipo](/javascript/api/excel/excel.refreshmodechangedeventargs#type)|Obtém o tipo do evento.|
|[RefreshRequestCompletedEventArgs](/javascript/api/excel/excel.refreshrequestcompletedeventargs)|[atualizado](/javascript/api/excel/excel.refreshrequestcompletedeventargs#refreshed)|Indica se a solicitação de atualização foi bem-sucedida.|
||[serviceId](/javascript/api/excel/excel.refreshrequestcompletedeventargs#serviceId)|A ID exclusiva do objeto cuja solicitação de atualização foi concluída.|
||[source](/javascript/api/excel/excel.refreshrequestcompletedeventargs#source)|Obtém a origem do evento.|
||[tipo](/javascript/api/excel/excel.refreshrequestcompletedeventargs#type)|Obtém o tipo do evento.|
||[avisos](/javascript/api/excel/excel.refreshrequestcompletedeventargs#warnings)|Uma matriz que contém quaisquer avisos gerados a partir da solicitação de atualização.|
|[Shape](/javascript/api/excel/excel.shape)|[displayName](/javascript/api/excel/excel.shape#displayName)|Obtém o nome de exibição da forma.|
|[ShapeCollection](/javascript/api/excel/excel.shapecollection)|[addSvg(xml: string)](/javascript/api/excel/excel.shapecollection#addSvg_xml_)|Cria um gráfico vetorial escalável (SVG) de uma cadeia de caracteres XML e a adiciona à planilha.|
||[getItemOrNullObject(key: string)](/javascript/api/excel/excel.shapecollection#getItemOrNullObject_key_)|Obtém uma forma usando seu nome ou ID.|
|[Segmentação de dados](/javascript/api/excel/excel.slicer)|[nameInFormula](/javascript/api/excel/excel.slicer#nameInFormula)|Representa o nome da segmentação de dados usada na fórmula.|
||[slicerStyle](/javascript/api/excel/excel.slicer#slicerStyle)|O estilo aplicado à slicer.|
||[setStyle(style: string \| SlicerStyle \| BuiltInSlicerStyle)](/javascript/api/excel/excel.slicer#setStyle_style_)|Define o estilo aplicado à slicer.|
|[SpillErrorCellValue](/javascript/api/excel/excel.spillerrorcellvalue)|[errorSubType](/javascript/api/excel/excel.spillerrorcellvalue#errorSubType)|Representa o tipo de `SpillErrorCellValue` .|
||[errorType](/javascript/api/excel/excel.spillerrorcellvalue#errorType)|Representa o tipo de `ErrorCellValue` .|
||[primitive](/javascript/api/excel/excel.spillerrorcellvalue#primitive)|Representa o valor que seria retornado por `Range.values` uma célula com esse valor.|
||[primitiveType](/javascript/api/excel/excel.spillerrorcellvalue#primitiveType)|Representa o valor que seria retornado por `Range.valueTypes` uma célula com esse valor.|
||[type](/javascript/api/excel/excel.spillerrorcellvalue#type)|Representa o tipo desse valor de célula.|
|[StringCellValue](/javascript/api/excel/excel.stringcellvalue)|[primitive](/javascript/api/excel/excel.stringcellvalue#primitive)|Representa o valor que seria retornado por `Range.values` uma célula com esse valor.|
||[primitiveType](/javascript/api/excel/excel.stringcellvalue#primitiveType)|Representa o valor que seria retornado por `Range.valueTypes` uma célula com esse valor.|
||[type](/javascript/api/excel/excel.stringcellvalue#type)|Representa o tipo desse valor de célula.|
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
|[ValueErrorCellValue](/javascript/api/excel/excel.valueerrorcellvalue)|[errorSubType](/javascript/api/excel/excel.valueerrorcellvalue#errorSubType)|Representa o tipo de `ValueErrorCellValue` .|
||[errorType](/javascript/api/excel/excel.valueerrorcellvalue#errorType)|Representa o tipo de `ErrorCellValue` .|
||[primitive](/javascript/api/excel/excel.valueerrorcellvalue#primitive)|Representa o valor que seria retornado por `Range.values` uma célula com esse valor.|
||[primitiveType](/javascript/api/excel/excel.valueerrorcellvalue#primitiveType)|Representa o valor que seria retornado por `Range.valueTypes` uma célula com esse valor.|
||[type](/javascript/api/excel/excel.valueerrorcellvalue#type)|Representa o tipo desse valor de célula.|
|[ValueTypeNotAvailableCellValue](/javascript/api/excel/excel.valuetypenotavailablecellvalue)|[primitive](/javascript/api/excel/excel.valuetypenotavailablecellvalue#primitive)|Representa o valor que seria retornado por `Range.values` uma célula com esse valor.|
||[primitiveType](/javascript/api/excel/excel.valuetypenotavailablecellvalue#primitiveType)|Representa o valor que seria retornado por `Range.valueTypes` uma célula com esse valor.|
||[type](/javascript/api/excel/excel.valuetypenotavailablecellvalue#type)|Representa o tipo desse valor de célula.|
|[WebImageCellValue](/javascript/api/excel/excel.webimagecellvalue)|[address](/javascript/api/excel/excel.webimagecellvalue#address)|Representa a URL da qual a imagem será baixada.|
||[altText](/javascript/api/excel/excel.webimagecellvalue#altText)|Representa o texto alternativo que pode ser usado em cenários de acessibilidade para descrever o que a imagem representa.|
||[attribution](/javascript/api/excel/excel.webimagecellvalue#attribution)|Representa informações de atribuição para descrever os requisitos de origem e licença para usar essa imagem.|
||[primitive](/javascript/api/excel/excel.webimagecellvalue#primitive)|Representa o valor que seria retornado por `Range.values` uma célula com esse valor.|
||[primitiveType](/javascript/api/excel/excel.webimagecellvalue#primitiveType)|Representa o valor que seria retornado por `Range.valueTypes` uma célula com esse valor.|
||[provider](/javascript/api/excel/excel.webimagecellvalue#provider)|Representa informações que descrevem a entidade ou indivíduo que forneceu a imagem.|
||[relatedImagesAddress](/javascript/api/excel/excel.webimagecellvalue#relatedImagesAddress)|Representa a URL de uma página da Web com imagens consideradas relacionadas a este `WebImageCellValue` .|
||[type](/javascript/api/excel/excel.webimagecellvalue#type)|Representa o tipo desse valor de célula.|
|[Workbook](/javascript/api/excel/excel.workbook)|[linkedDataTypes](/javascript/api/excel/excel.workbook#linkedDataTypes)|Retorna uma coleção de tipos de dados vinculados que fazem parte da lista de trabalho.|
||[consultas](/javascript/api/excel/excel.workbook#queries)|Retorna uma coleção de consultas do Power Query que fazem parte da workbook.|
||[tasks](/javascript/api/excel/excel.workbook#tasks)|Retorna uma coleção de tarefas que estão presentes na workbook.|
||[showPivotFieldList](/javascript/api/excel/excel.workbook#showPivotFieldList)|Especifica se o painel de lista de campos da Tabela Dinâmica é mostrado no nível da lista de trabalho.|
||[use1904DateSystem](/javascript/api/excel/excel.workbook#use1904DateSystem)|True se a pasta de trabalho usar o sistema de dados 1904.|
|[Worksheet](/javascript/api/excel/excel.worksheet)|[onFiltered](/javascript/api/excel/excel.worksheet#onFiltered)|Ocorre quando um filtro é aplicado em uma planilha específica.|
||[onNameChanged](/javascript/api/excel/excel.worksheet#onNameChanged)|Ocorre quando o nome da planilha é alterado.|
||[onProtectionChanged](/javascript/api/excel/excel.worksheet#onProtectionChanged)|Ocorre quando o estado de proteção da planilha é alterado.|
||[onVisibilityChanged](/javascript/api/excel/excel.worksheet#onVisibilityChanged)|Ocorre quando a visibilidade da planilha é alterada.|
||[tabId](/javascript/api/excel/excel.worksheet#tabId)|Retorna um valor que representa essa planilha que pode ser lido por Open Office XML.|
||[tasks](/javascript/api/excel/excel.worksheet#tasks)|Retorna uma coleção de tarefas presentes na planilha.|
|[WorksheetChangedEventArgs](/javascript/api/excel/excel.worksheetchangedeventargs)|[changeDirectionState](/javascript/api/excel/excel.worksheetchangedeventargs#changeDirectionState)|Representa uma alteração na direção em que as células de uma planilha serão deslocadas quando uma célula ou células são excluídas ou inseridas.|
||[triggerSource](/javascript/api/excel/excel.worksheetchangedeventargs#triggerSource)|Representa a origem do gatilho do evento.|
|[WorksheetCollection](/javascript/api/excel/excel.worksheetcollection)|[addFromBase64(base64File: string, sheetNamesToInsert?: string[], positionType?: Excel.WorksheetPositionType, relativeTo?: Worksheet \| string)](/javascript/api/excel/excel.worksheetcollection#addFromBase64_base64File__sheetNamesToInsert__positionType__relativeTo_)|Insere as planilhas especificadas de uma pasta de trabalho na pasta de trabalho atual.|
||[onFiltered](/javascript/api/excel/excel.worksheetcollection#onFiltered)|Ocorre quando filtro de uma planilha é aplicado na pasta de trabalho.|
||[onMoved](/javascript/api/excel/excel.worksheetcollection#onMoved)|Ocorre quando uma planilha é movida por um usuário dentro de uma pasta de trabalho.|
||[onNameChanged](/javascript/api/excel/excel.worksheetcollection#onNameChanged)|Ocorre quando o nome da planilha é alterado na coleção de planilhas.|
||[onProtectionChanged](/javascript/api/excel/excel.worksheetcollection#onProtectionChanged)|Ocorre quando o estado de proteção da planilha é alterado.|
||[onVisibilityChanged](/javascript/api/excel/excel.worksheetcollection#onVisibilityChanged)|Ocorre quando a visibilidade da planilha é alterada na coleção de planilhas.|
|[WorksheetFilteredEventArgs](/javascript/api/excel/excel.worksheetfilteredeventargs)|[tipo](/javascript/api/excel/excel.worksheetfilteredeventargs#type)|Obtém o tipo do evento.|
||[worksheetId](/javascript/api/excel/excel.worksheetfilteredeventargs#worksheetId)|Obtém a ID da planilha na qual o filtro é aplicado.|
|[WorksheetMovedEventArgs](/javascript/api/excel/excel.worksheetmovedeventargs)|[positionAfter](/javascript/api/excel/excel.worksheetmovedeventargs#positionAfter)|Obtém a nova posição da planilha, após a movimentação.|
||[positionBefore](/javascript/api/excel/excel.worksheetmovedeventargs#positionBefore)|Obtém a posição anterior da planilha, antes da movimentação.|
||[source](/javascript/api/excel/excel.worksheetmovedeventargs#source)|A origem do evento.|
||[tipo](/javascript/api/excel/excel.worksheetmovedeventargs#type)|Obtém o tipo do evento.|
||[worksheetId](/javascript/api/excel/excel.worksheetmovedeventargs#worksheetId)|Obtém a ID da planilha que foi movida.|
|[WorksheetNameChangedEventArgs](/javascript/api/excel/excel.worksheetnamechangedeventargs)|[nameAfter](/javascript/api/excel/excel.worksheetnamechangedeventargs#nameAfter)|Obtém o novo nome da planilha, após a alteração do nome.|
||[nameBefore](/javascript/api/excel/excel.worksheetnamechangedeventargs#nameBefore)|Obtém o nome anterior da planilha, antes de o nome ser alterado.|
||[source](/javascript/api/excel/excel.worksheetnamechangedeventargs#source)|A origem do evento.|
||[tipo](/javascript/api/excel/excel.worksheetnamechangedeventargs#type)|Obtém o tipo do evento.|
||[worksheetId](/javascript/api/excel/excel.worksheetnamechangedeventargs#worksheetId)|Obtém a ID da planilha com o novo nome.|
|[WorksheetProtection](/javascript/api/excel/excel.worksheetprotection)|[checkPassword(password?: string)](/javascript/api/excel/excel.worksheetprotection#checkPassword_password_)|Especifica se a senha pode ser usada para desbloquear a proteção da planilha.|
||[pauseProtection(password?: string)](/javascript/api/excel/excel.worksheetprotection#pauseProtection_password_)|Pausa a proteção da planilha para o objeto de planilha determinado para o usuário em uma determinada sessão.|
||[allowEditRanges](/javascript/api/excel/excel.worksheetprotection#allowEditRanges)|Especifica o `AllowEditRangeCollection` encontrado nesta planilha.|
||[canPauseProtection](/javascript/api/excel/excel.worksheetprotection#canPauseProtection)|Especifica se a proteção pode ser pausada para esta planilha.|
||[isPasswordProtected](/javascript/api/excel/excel.worksheetprotection#isPasswordProtected)|Especifica se a planilha está protegida por senha.|
||[isPaused](/javascript/api/excel/excel.worksheetprotection#isPaused)|Especifica se a proteção da planilha está pausada.|
||[resumeProtection()](/javascript/api/excel/excel.worksheetprotection#resumeProtection__)|Retoma a proteção da planilha para o objeto de planilha determinado para o usuário em uma determinada sessão.|
||[setPassword(password?: string)](/javascript/api/excel/excel.worksheetprotection#setPassword_password_)|Altera a senha associada ao `WorksheetProtection` objeto.|
||[updateOptions(options: Excel. WorksheetProtectionOptions)](/javascript/api/excel/excel.worksheetprotection#updateOptions_options_)|Altere as opções de proteção da planilha associadas ao `WorksheetProtection` objeto.|
|[WorksheetProtectionChangedEventArgs](/javascript/api/excel/excel.worksheetprotectionchangedeventargs)|[allowEditRangesChanged](/javascript/api/excel/excel.worksheetprotectionchangedeventargs#allowEditRangesChanged)|Especifica se algum dos `AllowEditRange` objetos foi alterado.|
||[isProtected](/javascript/api/excel/excel.worksheetprotectionchangedeventargs#isProtected)|Obtém o status de proteção atual da planilha.|
||[protectionOptionsChanged](/javascript/api/excel/excel.worksheetprotectionchangedeventargs#protectionOptionsChanged)|Especifica se o `WorksheetProtectionOptions` foi alterado.|
||[sheetPasswordChanged](/javascript/api/excel/excel.worksheetprotectionchangedeventargs#sheetPasswordChanged)|Especifica se a senha da planilha foi alterada.|
||[source](/javascript/api/excel/excel.worksheetprotectionchangedeventargs#source)|A origem do evento.|
||[tipo](/javascript/api/excel/excel.worksheetprotectionchangedeventargs#type)|Obtém o tipo do evento.|
||[worksheetId](/javascript/api/excel/excel.worksheetprotectionchangedeventargs#worksheetId)|Obtém a ID da planilha na qual o status da proteção é alterado.|
|[WorksheetVisibilityChangedEventArgs](/javascript/api/excel/excel.worksheetvisibilitychangedeventargs)|[source](/javascript/api/excel/excel.worksheetvisibilitychangedeventargs#source)|A origem do evento.|
||[tipo](/javascript/api/excel/excel.worksheetvisibilitychangedeventargs#type)|Obtém o tipo do evento.|
||[visibilityAfter](/javascript/api/excel/excel.worksheetvisibilitychangedeventargs#visibilityAfter)|Obtém a nova configuração de visibilidade da planilha, após a alteração de visibilidade.|
||[visibilityBefore](/javascript/api/excel/excel.worksheetvisibilitychangedeventargs#visibilityBefore)|Obtém a configuração de visibilidade anterior da planilha, antes da alteração de visibilidade.|
||[worksheetId](/javascript/api/excel/excel.worksheetvisibilitychangedeventargs#worksheetId)|Obtém a ID da planilha cuja visibilidade foi alterada.|

## <a name="see-also"></a>Confira também

- [Documentação deReferência da API JavaScript do Excel](/javascript/api/excel?view=excel-js-preview&preserve-view=true)
- [Conjuntos de requisitos da API JavaScript do Excel](excel-api-requirement-sets.md)
