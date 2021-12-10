---
title: APIs de visualização do JavaScript para Excel
description: Detalhes sobre as próximas Excel APIs JavaScript.
ms.date: 12/08/2021
ms.prod: excel
ms.localizationpriority: medium
ms.openlocfilehash: 32a2f5d355086c51cbf165dd7ed335e96c96647a
ms.sourcegitcommit: ddb1d85186fd6e77d732159430d20eb7395b9a33
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 12/10/2021
ms.locfileid: "61406638"
---
# <a name="excel-javascript-preview-apis"></a>APIs de visualização do JavaScript para Excel

As novas APIs do JavaScript para Excel são introduzidas pela primeira vez na "visualização" e, posteriormente, tornam-se parte de um conjunto específico de requisitos numerados, após passarem por vários testes e após a recolha das opiniões de usuários.

[!INCLUDE [Information about using preview APIs](../../includes/using-preview-apis-host.md)]

A tabela a seguir fornece um resumo conciso das APIs, enquanto a tabela de lista [de API](#api-list) subsequente fornece uma lista detalhada.

| Área de recurso | Descrição | Objetos relevantes |
|:--- |:--- |:--- |
| [Tipos de dados](../../excel/excel-data-types-overview.md) | Uma extensão de tipos de dados Excel existentes, incluindo suporte para números formatados e imagens da Web. | [ArrayCellValue](/javascript/api/excel/excel.arraycellvalue), [BooleanCellValue](/javascript/api/excel/excel.booleancellvalue), [CellValueAttributionAttributes](/javascript/api/excel/excel.cellvalueattributionattributes), [CellValueProviderAttributes](/javascript/api/excel/excel.cellvalueproviderattributes), [DoubleCellValue](/javascript/api/excel/excel.doublecellvalue), [EmptyCellValue](/javascript/api/excel/excel.emptycellvalue), [EntityCellValue](/javascript/api/excel/excel.entitycellvalue), [FormattedNumberCellValue](/javascript/api/excel/excel.formattednumbercellvalue), [StringCellValue](/javascript/api/excel/excel.stringcellvalue), [ValueTypeNotAvailableCellValue](/javascript/api/excel/excel.valuetypenotavailablecellvalue), [WebImageCellValue](/javascript/api/excel/excel.webimagecellvalue) |
| [Erros de tipos de dados](../../excel/excel-data-types-concepts.md#improved-error-support) | Objetos de erro que suportam tipos de dados expandidos. | [BlockedErrorCellValue](/javascript/api/excel/excel.blockederrorcellvalue), [BusyErrorCellValue](/javascript/api/excel/excel.busyerrorcellvalue), [CalcErrorCellValue](/javascript/api/excel/excel.calcerrorcellvalue), [ConnectErrorCellValue](/javascript/api/excel/excel.connecterrorcellvalue), [Div0ErrorCellValue](/javascript/api/excel/excel.div0errorcellvalue), [FieldErrorCellValue](/javascript/api/excel/excel.fielderrorcellvalue), [GettingDataErrorCellValue](/javascript/api/excel/excel.gettingdataerrorcellvalue), [NotAvailableErrorCellValue](/javascript/api/excel/excel.notavailableerrorcellvalue), [NameErrorCellValue](/javascript/api/excel/excel.nameerrorcellvalue), [NullErrorCellValue](/javascript/api/excel/excel.nullerrorcellvalue), [NumErrorCellValue](/javascript/api/excel/excel.numerrorcellvalue), [RefErrorCellValue](/javascript/api/excel/excel.referrorcellvalue), [SpillErrorCellValue](/javascript/api/excel/excel.spillerrorcellvalue), [ValueErrorCellValue](/javascript/api/excel/excel.valueerrorcellvalue)|
| Tarefas do documento | Transforme os comentários em tarefas atribuídas aos usuários. | [DocumentTask](/javascript/api/excel/excel.documenttask) |
| Identidades | Gerencie identidades de usuário, incluindo nome de exibição e endereço de email. | [Identity](/javascript/api/excel/excel.identity), [IdentityCollection](/javascript/api/excel/excel.identitycollection), [IdentityEntity](/javascript/api/excel/excel.identityentity) |
| Tipos de dados vinculados | Adiciona suporte para tipos de dados conectados Excel de fontes externas. | [LinkedDataType](/javascript/api/excel/excel.linkeddatatype), [LinkedDataTypeAddedEventArgs](/javascript/api/excel/excel.linkeddatatypeaddedeventargs), [LinkedDataTypeCollection](/javascript/api/excel/excel.linkeddatatypecollection) |
| Table styles | Fornece controle para fonte, borda, cor de preenchimento e outros aspectos dos estilos de tabela. | [Tabela,](/javascript/api/excel/excel.table) [Tabela Dinâmica,](/javascript/api/excel/excel.pivottable) [Slicer](/javascript/api/excel/excel.slicer) |
| Proteção de planilha | Impedir que usuários não autorizados mudem para intervalos especificados em uma planilha. | [WorksheetProtection](/javascript/api/excel/excel.worksheetprotection), [AllowEditRange](/javascript/api/excel/excel.alloweditrange), [AllowEditRangeCollection](/javascript/api/excel/excel.alloweditrangecollection), [AllowEditRangeOptions](/javascript/api/excel/excel.alloweditrangeoptions) |

## <a name="api-list"></a>Lista de API

A tabela a seguir lista as Excel APIs JavaScript atualmente em visualização. Para uma lista completa de todas as EXCEL JavaScript (incluindo APIs de visualização e APIs lançadas anteriormente), consulte todas as [APIs JavaScript](/javascript/api/excel?view=excel-js-preview&preserve-view=true)Excel JavaScript .

| Classe | Campos | Descrição |
|:---|:---|:---|
|[AllowEditRange](/javascript/api/excel/excel.alloweditrange)|[address](/javascript/api/excel/excel.alloweditrange#address)|Especifica o intervalo associado ao objeto.|
||[delete()](/javascript/api/excel/excel.alloweditrange#delete__)|Exclui esse objeto do `AllowEditRangeCollection` .|
||[isPasswordProtected](/javascript/api/excel/excel.alloweditrange#isPasswordProtected)|Especifica se a `AllowEditRange` senha está protegida.|
||[pauseProtection(password?: string)](/javascript/api/excel/excel.alloweditrange#pauseProtection_password_)|Pausa a proteção da planilha para o `AllowEditRange` objeto dado para o usuário em uma determinada sessão.|
||[setPassword(password?: string)](/javascript/api/excel/excel.alloweditrange#setPassword_password_)|Altera a senha associada ao `AllowEditRange` .|
||[title](/javascript/api/excel/excel.alloweditrange#title)|Especifica o título do objeto.|
|[AllowEditRangeCollection](/javascript/api/excel/excel.alloweditrangecollection)|[add(title: string, rangeAddress: string, options?: Excel. AllowEditRangeOptions)](/javascript/api/excel/excel.alloweditrangecollection#add_title__rangeAddress__options_)|Adiciona um `AllowEditRange` objeto à coleção.|
||[getCount()](/javascript/api/excel/excel.alloweditrangecollection#getCount__)|Retorna o número `AllowEditRange` de objetos na coleção.|
||[getItem(key: string)](/javascript/api/excel/excel.alloweditrangecollection#getItem_key_)|Obtém `AllowEditRange` o objeto pelo título.|
||[getItemAt(index: number)](/javascript/api/excel/excel.alloweditrangecollection#getItemAt_index_)|Retorna um `AllowEditRange` objeto pelo índice na coleção.|
||[getItemOrNullObject(key: string)](/javascript/api/excel/excel.alloweditrangecollection#getItemOrNullObject_key_)|Obtém `AllowEditRange` o objeto pelo título.|
||[items](/javascript/api/excel/excel.alloweditrangecollection#items)|Obtém os itens filhos carregados nesta coleção.|
||[pauseProtection(password: string)](/javascript/api/excel/excel.alloweditrangecollection#pauseProtection_password_)|Pausa a proteção da planilha para todos os objetos da coleção que têm a `AllowEditRange` senha dada para o usuário em uma determinada sessão.|
|[AllowEditRangeOptions](/javascript/api/excel/excel.alloweditrangeoptions)|[senha](/javascript/api/excel/excel.alloweditrangeoptions#password)|A senha associada ao `AllowEditRange` .|
|[ArrayCellValue](/javascript/api/excel/excel.arraycellvalue)|[basicType](/javascript/api/excel/excel.arraycellvalue#basicType)|Representa o valor que seria retornado por `Range.valueTypes` uma célula com esse valor.|
||[basicValue](/javascript/api/excel/excel.arraycellvalue#basicValue)|Representa o valor que seria retornado por `Range.values` uma célula com esse valor.|
||[elements](/javascript/api/excel/excel.arraycellvalue#elements)|Representa os elementos da matriz.|
||[tipo](/javascript/api/excel/excel.arraycellvalue#type)|Representa o tipo desse valor de célula.|
|[BlockedErrorCellValue](/javascript/api/excel/excel.blockederrorcellvalue)|[basicType](/javascript/api/excel/excel.blockederrorcellvalue#basicType)|Representa o valor que seria retornado por `Range.valueTypes` uma célula com esse valor.|
||[basicValue](/javascript/api/excel/excel.blockederrorcellvalue#basicValue)|Representa o valor que seria retornado por `Range.values` uma célula com esse valor.|
||[errorSubType](/javascript/api/excel/excel.blockederrorcellvalue#errorSubType)|Representa o tipo de `BlockedErrorCellValue` .|
||[errorType](/javascript/api/excel/excel.blockederrorcellvalue#errorType)|Representa o tipo de `ErrorCellValue` .|
||[tipo](/javascript/api/excel/excel.blockederrorcellvalue#type)|Representa o tipo desse valor de célula.|
|[BooleanCellValue](/javascript/api/excel/excel.booleancellvalue)|[basicType](/javascript/api/excel/excel.booleancellvalue#basicType)|Representa o valor que seria retornado por `Range.valueTypes` uma célula com esse valor.|
||[basicValue](/javascript/api/excel/excel.booleancellvalue#basicValue)|Representa o valor que seria retornado por `Range.values` uma célula com esse valor.|
||[tipo](/javascript/api/excel/excel.booleancellvalue#type)|Representa o tipo desse valor de célula.|
|[BusyErrorCellValue](/javascript/api/excel/excel.busyerrorcellvalue)|[basicType](/javascript/api/excel/excel.busyerrorcellvalue#basicType)|Representa o valor que seria retornado por `Range.valueTypes` uma célula com esse valor.|
||[basicValue](/javascript/api/excel/excel.busyerrorcellvalue#basicValue)|Representa o valor que seria retornado por `Range.values` uma célula com esse valor.|
||[errorSubType](/javascript/api/excel/excel.busyerrorcellvalue#errorSubType)|Representa o tipo de `BusyErrorCellValue` .|
||[errorType](/javascript/api/excel/excel.busyerrorcellvalue#errorType)|Representa o tipo de `ErrorCellValue` .|
||[tipo](/javascript/api/excel/excel.busyerrorcellvalue#type)|Representa o tipo desse valor de célula.|
|[CalcErrorCellValue](/javascript/api/excel/excel.calcerrorcellvalue)|[basicType](/javascript/api/excel/excel.calcerrorcellvalue#basicType)|Representa o valor que seria retornado por `Range.valueTypes` uma célula com esse valor.|
||[basicValue](/javascript/api/excel/excel.calcerrorcellvalue#basicValue)|Representa o valor que seria retornado por `Range.values` uma célula com esse valor.|
||[errorSubType](/javascript/api/excel/excel.calcerrorcellvalue#errorSubType)|Representa o tipo de `CalcErrorCellValue` .|
||[errorType](/javascript/api/excel/excel.calcerrorcellvalue#errorType)|Representa o tipo de `ErrorCellValue` .|
||[tipo](/javascript/api/excel/excel.calcerrorcellvalue#type)|Representa o tipo desse valor de célula.|
|[CardLayoutListSection](/javascript/api/excel/excel.cardlayoutlistsection)|[layout](/javascript/api/excel/excel.cardlayoutlistsection#layout)|Representa o tipo de layout desta seção.|
|[CardLayoutPropertyReference](/javascript/api/excel/excel.cardlayoutpropertyreference)|[property](/javascript/api/excel/excel.cardlayoutpropertyreference#property)|O nome da propriedade referenciada pelo layout do cartão.|
|[CardLayoutSectionStandardProperties](/javascript/api/excel/excel.cardlayoutsectionstandardproperties)|[collapsed](/javascript/api/excel/excel.cardlayoutsectionstandardproperties#collapsed)|Representa se esta seção do cartão está inicialmente recolhido.|
||[collapsible](/javascript/api/excel/excel.cardlayoutsectionstandardproperties#collapsible)|Representa se essa seção do cartão é retutível.|
||[properties](/javascript/api/excel/excel.cardlayoutsectionstandardproperties#properties)|Representa os nomes das propriedades nesta seção.|
||[title](/javascript/api/excel/excel.cardlayoutsectionstandardproperties#title)|Representa o título desta seção do cartão.|
|[CardLayoutStandardProperties](/javascript/api/excel/excel.cardlayoutstandardproperties)|[mainImage](/javascript/api/excel/excel.cardlayoutstandardproperties#mainImage)|Especifica uma propriedade que será usada como a imagem principal do cartão.|
||[sections](/javascript/api/excel/excel.cardlayoutstandardproperties#sections)|Representa as seções do cartão.|
||[subTitle](/javascript/api/excel/excel.cardlayoutstandardproperties#subTitle)|Representa uma especificação de qual propriedade contém o subtítulo do cartão.|
||[title](/javascript/api/excel/excel.cardlayoutstandardproperties#title)|Representa o título do cartão ou a especificação de qual propriedade contém o título do cartão.|
|[CardLayoutTableSection](/javascript/api/excel/excel.cardlayouttablesection)|[layout](/javascript/api/excel/excel.cardlayouttablesection#layout)|Representa o tipo de layout desta seção.|
|[CellValueAttributionAttributes](/javascript/api/excel/excel.cellvalueattributionattributes)|[licenseAddress](/javascript/api/excel/excel.cellvalueattributionattributes#licenseAddress)|Representa uma URL para uma licença ou fonte que descreve como essa propriedade pode ser usada.|
||[licenseText](/javascript/api/excel/excel.cellvalueattributionattributes#licenseText)|Representa um nome para a licença que rege essa propriedade.|
||[sourceAddress](/javascript/api/excel/excel.cellvalueattributionattributes#sourceAddress)|Representa uma URL para a origem do `CellValue` .|
||[sourceText](/javascript/api/excel/excel.cellvalueattributionattributes#sourceText)|Representa um nome para a origem do `CellValue` .|
|[CellValuePropertyMetadata](/javascript/api/excel/excel.cellvaluepropertymetadata)|[attribution](/javascript/api/excel/excel.cellvaluepropertymetadata#attribution)|Representa informações de atribuição para descrever os requisitos de origem e licença para usar essa propriedade.|
||[excludeFrom](/javascript/api/excel/excel.cellvaluepropertymetadata#excludeFrom)|Representa de quais recursos essa propriedade é excluída.|
||[sub-rótulo](/javascript/api/excel/excel.cellvaluepropertymetadata#sublabel)|Representa o sub-rótulo dessa propriedade mostrado no exibição de cartão.|
|[CellValuePropertyMetadataExclusions](/javascript/api/excel/excel.cellvaluepropertymetadataexclusions)|[autoComplete](/javascript/api/excel/excel.cellvaluepropertymetadataexclusions#autoComplete)|True representa que a propriedade é excluída das propriedades mostradas pela conclusão automática.|
||[calcCompare](/javascript/api/excel/excel.cellvaluepropertymetadataexclusions#calcCompare)|True representa que a propriedade é excluída das propriedades usadas para comparar valores de células durante o recalcal.|
||[cardView](/javascript/api/excel/excel.cellvaluepropertymetadataexclusions#cardView)|True representa que a propriedade é excluída das propriedades mostradas pelo exibição de cartão.|
||[dotNotation](/javascript/api/excel/excel.cellvaluepropertymetadataexclusions#dotNotation)|True representa que a propriedade é excluída das propriedades que podem ser acessadas por meio da função FIELDVALUE.|
|[CellValueProviderAttributes](/javascript/api/excel/excel.cellvalueproviderattributes)|[description](/javascript/api/excel/excel.cellvalueproviderattributes#description)|Representa a propriedade de descrição do provedor usada no exibição de cartão se nenhum logotipo for especificado.|
||[logoSourceAddress](/javascript/api/excel/excel.cellvalueproviderattributes#logoSourceAddress)|Representa uma URL usada para baixar uma imagem que será usada como um logotipo no exibição de cartão.|
||[logoTargetAddress](/javascript/api/excel/excel.cellvalueproviderattributes#logoTargetAddress)|Representa uma URL que será o destino de navegação se o usuário clicar no elemento logo no modo de exibição de cartão.|
|[Comentário](/javascript/api/excel/excel.comment)|[assignTask(assignee: Identity)](/javascript/api/excel/excel.comment#assignTask_assignee_)|Atribui a tarefa anexada ao comentário ao usuário dado como um destinatário.|
||[getTask()](/javascript/api/excel/excel.comment#getTask__)|Obtém a tarefa associada a este comentário.|
||[getTaskOrNullObject()](/javascript/api/excel/excel.comment#getTaskOrNullObject__)|Obtém a tarefa associada a este comentário.|
|[CommentReply](/javascript/api/excel/excel.commentreply)|[assignTask(assignee: Identity)](/javascript/api/excel/excel.commentreply#assignTask_assignee_)|Atribui a tarefa anexada ao comentário ao usuário determinado como o único destinatário.|
||[getTask()](/javascript/api/excel/excel.commentreply#getTask__)|Obtém a tarefa associada ao thread desta resposta de comentário.|
||[getTaskOrNullObject()](/javascript/api/excel/excel.commentreply#getTaskOrNullObject__)|Obtém a tarefa associada ao thread desta resposta de comentário.|
|[ConnectErrorCellValue](/javascript/api/excel/excel.connecterrorcellvalue)|[basicType](/javascript/api/excel/excel.connecterrorcellvalue#basicType)|Representa o valor que seria retornado por `Range.valueTypes` uma célula com esse valor.|
||[basicValue](/javascript/api/excel/excel.connecterrorcellvalue#basicValue)|Representa o valor que seria retornado por `Range.values` uma célula com esse valor.|
||[errorSubType](/javascript/api/excel/excel.connecterrorcellvalue#errorSubType)|Representa o tipo de `ConnectErrorCellValue` .|
||[errorType](/javascript/api/excel/excel.connecterrorcellvalue#errorType)|Representa o tipo de `ErrorCellValue` .|
||[tipo](/javascript/api/excel/excel.connecterrorcellvalue#type)|Representa o tipo desse valor de célula.|
|[Div0ErrorCellValue](/javascript/api/excel/excel.div0errorcellvalue)|[basicType](/javascript/api/excel/excel.div0errorcellvalue#basicType)|Representa o valor que seria retornado por `Range.valueTypes` uma célula com esse valor.|
||[basicValue](/javascript/api/excel/excel.div0errorcellvalue#basicValue)|Representa o valor que seria retornado por `Range.values` uma célula com esse valor.|
||[errorType](/javascript/api/excel/excel.div0errorcellvalue#errorType)|Representa o tipo de `ErrorCellValue` .|
||[tipo](/javascript/api/excel/excel.div0errorcellvalue#type)|Representa o tipo desse valor de célula.|
|[DocumentTask](/javascript/api/excel/excel.documenttask)|[assignees](/javascript/api/excel/excel.documenttask#assignees)|Retorna uma coleção de atribuídos da tarefa.|
||[changes](/javascript/api/excel/excel.documenttask#changes)|Obtém os registros de alteração da tarefa.|
||[comment](/javascript/api/excel/excel.documenttask#comment)|Obtém o comentário associado à tarefa.|
||[completedBy](/javascript/api/excel/excel.documenttask#completedBy)|Obtém o usuário mais recente para ter concluído a tarefa.|
||[completedDateTime](/javascript/api/excel/excel.documenttask#completedDateTime)|Obtém a data e a hora em que a tarefa foi concluída.|
||[createdBy](/javascript/api/excel/excel.documenttask#createdBy)|Obtém o usuário que criou a tarefa.|
||[createdDateTime](/javascript/api/excel/excel.documenttask#createdDateTime)|Obtém a data e a hora em que a tarefa foi criada.|
||[id](/javascript/api/excel/excel.documenttask#id)|Obtém a ID da tarefa.|
||[percentComplete](/javascript/api/excel/excel.documenttask#percentComplete)|Especifica a porcentagem de conclusão da tarefa.|
||[priority](/javascript/api/excel/excel.documenttask#priority)|Especifica a prioridade da tarefa.|
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
||[tipo](/javascript/api/excel/excel.documenttaskchange#type)|Representa o tipo de ação do registro de alteração de tarefa.|
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
|[DoubleCellValue](/javascript/api/excel/excel.doublecellvalue)|[basicType](/javascript/api/excel/excel.doublecellvalue#basicType)|Representa o valor que seria retornado por `Range.valueTypes` uma célula com esse valor.|
||[basicValue](/javascript/api/excel/excel.doublecellvalue#basicValue)|Representa o valor que seria retornado por `Range.values` uma célula com esse valor.|
||[tipo](/javascript/api/excel/excel.doublecellvalue#type)|Representa o tipo desse valor de célula.|
|[EmptyCellValue](/javascript/api/excel/excel.emptycellvalue)|[basicType](/javascript/api/excel/excel.emptycellvalue#basicType)|Representa o valor que seria retornado por `Range.valueTypes` uma célula com esse valor.|
||[basicValue](/javascript/api/excel/excel.emptycellvalue#basicValue)|Representa o valor que seria retornado por `Range.values` uma célula com esse valor.|
||[tipo](/javascript/api/excel/excel.emptycellvalue#type)|Representa o tipo desse valor de célula.|
|[EntityCardLayout](/javascript/api/excel/excel.entitycardlayout)|[layout](/javascript/api/excel/excel.entitycardlayout#layout)|Representa o tipo desse layout.|
|[EntityCellValue](/javascript/api/excel/excel.entitycellvalue)|[basicType](/javascript/api/excel/excel.entitycellvalue#basicType)|Representa o valor que seria retornado por `Range.valueTypes` uma célula com esse valor.|
||[basicValue](/javascript/api/excel/excel.entitycellvalue#basicValue)|Representa o valor que seria retornado por `Range.values` uma célula com esse valor.|
||[cardLayout](/javascript/api/excel/excel.entitycellvalue#cardLayout)|Representa o layout dessa entidade no exibição de cartão.|
||[properties: { [key: string]](/javascript/api/excel/excel.entitycellvalue#properties)|Representa as propriedades dessa entidade e seus metadados.|
||[text](/javascript/api/excel/excel.entitycellvalue#text)|Representa o texto mostrado quando uma célula com esse valor é renderizada.|
||[tipo](/javascript/api/excel/excel.entitycellvalue#type)|Representa o tipo desse valor de célula.|
|[FieldErrorCellValue](/javascript/api/excel/excel.fielderrorcellvalue)|[basicType](/javascript/api/excel/excel.fielderrorcellvalue#basicType)|Representa o valor que seria retornado por `Range.valueTypes` uma célula com esse valor.|
||[basicValue](/javascript/api/excel/excel.fielderrorcellvalue#basicValue)|Representa o valor que seria retornado por `Range.values` uma célula com esse valor.|
||[errorSubType](/javascript/api/excel/excel.fielderrorcellvalue#errorSubType)|Representa o tipo de `FieldErrorCellValue` .|
||[errorType](/javascript/api/excel/excel.fielderrorcellvalue#errorType)|Representa o tipo de `ErrorCellValue` .|
||[tipo](/javascript/api/excel/excel.fielderrorcellvalue#type)|Representa o tipo desse valor de célula.|
|[FormattedNumberCellValue](/javascript/api/excel/excel.formattednumbercellvalue)|[basicType](/javascript/api/excel/excel.formattednumbercellvalue#basicType)|Representa o valor que seria retornado por `Range.valueTypes` uma célula com esse valor.|
||[basicValue](/javascript/api/excel/excel.formattednumbercellvalue#basicValue)|Representa o valor que seria retornado por `Range.values` uma célula com esse valor.|
||[numberFormat](/javascript/api/excel/excel.formattednumbercellvalue#numberFormat)|Retorna a cadeia de caracteres de formato de número usada para exibir esse valor.|
||[tipo](/javascript/api/excel/excel.formattednumbercellvalue#type)|Representa o tipo desse valor de célula.|
|[GettingDataErrorCellValue](/javascript/api/excel/excel.gettingdataerrorcellvalue)|[basicType](/javascript/api/excel/excel.gettingdataerrorcellvalue#basicType)|Representa o valor que seria retornado por `Range.valueTypes` uma célula com esse valor.|
||[basicValue](/javascript/api/excel/excel.gettingdataerrorcellvalue#basicValue)|Representa o valor que seria retornado por `Range.values` uma célula com esse valor.|
||[errorType](/javascript/api/excel/excel.gettingdataerrorcellvalue#errorType)|Representa o tipo de `ErrorCellValue` .|
||[tipo](/javascript/api/excel/excel.gettingdataerrorcellvalue#type)|Representa o tipo desse valor de célula.|
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
||[requestRefresh()](/javascript/api/excel/excel.linkeddatatype#requestRefresh__)|Faz uma solicitação para atualizar o tipo de dados vinculado.|
||[requestSetRefreshMode(refreshMode: Excel. LinkedDataTypeRefreshMode)](/javascript/api/excel/excel.linkeddatatype#requestSetRefreshMode_refreshMode_)|Faz uma solicitação para alterar o modo de atualização para esse tipo de dados vinculado.|
||[serviceId](/javascript/api/excel/excel.linkeddatatype#serviceId)|A ID exclusiva do tipo de dados vinculado.|
||[supportedRefreshModes](/javascript/api/excel/excel.linkeddatatype#supportedRefreshModes)|Retorna uma matriz com todos os modos de atualização suportados pelo tipo de dados vinculado.|
|[LinkedDataTypeAddedEventArgs](/javascript/api/excel/excel.linkeddatatypeaddedeventargs)|[serviceId](/javascript/api/excel/excel.linkeddatatypeaddedeventargs#serviceId)|A ID exclusiva do novo tipo de dados vinculado.|
||[source](/javascript/api/excel/excel.linkeddatatypeaddedeventargs#source)|Obtém a origem do evento.|
||[tipo](/javascript/api/excel/excel.linkeddatatypeaddedeventargs#type)|Obtém o tipo do evento.|
|[LinkedDataTypeCollection](/javascript/api/excel/excel.linkeddatatypecollection)|[getCount()](/javascript/api/excel/excel.linkeddatatypecollection#getCount__)|Obtém o número de tipos de dados vinculados na coleção.|
||[getItem(key: number)](/javascript/api/excel/excel.linkeddatatypecollection#getItem_key_)|Obtém um tipo de dados vinculado por ID de serviço.|
||[getItemAt(index: number)](/javascript/api/excel/excel.linkeddatatypecollection#getItemAt_index_)|Obtém um tipo de dados vinculado pelo índice na coleção.|
||[getItemOrNullObject(key: number)](/javascript/api/excel/excel.linkeddatatypecollection#getItemOrNullObject_key_)|Obtém um tipo de dados vinculado por ID.|
||[items](/javascript/api/excel/excel.linkeddatatypecollection#items)|Obtém os itens filhos carregados nesta coleção.|
||[requestRefreshAll()](/javascript/api/excel/excel.linkeddatatypecollection#requestRefreshAll__)|Faz uma solicitação para atualizar todos os tipos de dados vinculados na coleção.|
|[NameErrorCellValue](/javascript/api/excel/excel.nameerrorcellvalue)|[basicType](/javascript/api/excel/excel.nameerrorcellvalue#basicType)|Representa o valor que seria retornado por `Range.valueTypes` uma célula com esse valor.|
||[basicValue](/javascript/api/excel/excel.nameerrorcellvalue#basicValue)|Representa o valor que seria retornado por `Range.values` uma célula com esse valor.|
||[errorType](/javascript/api/excel/excel.nameerrorcellvalue#errorType)|Representa o tipo de `ErrorCellValue` .|
||[tipo](/javascript/api/excel/excel.nameerrorcellvalue#type)|Representa o tipo desse valor de célula.|
|[NamedItem](/javascript/api/excel/excel.nameditem)|[valueAsJson](/javascript/api/excel/excel.nameditem#valueAsJson)|Uma representação JSON dos valores neste item nomeado.|
|[NamedItemArrayValues](/javascript/api/excel/excel.nameditemarrayvalues)|[valuesAsJson](/javascript/api/excel/excel.nameditemarrayvalues#valuesAsJson)|Uma representação JSON dos valores nas células nesse intervalo.|
|[NamedSheetViewCollection](/javascript/api/excel/excel.namedsheetviewcollection)|[getItemOrNullObject(key: string)](/javascript/api/excel/excel.namedsheetviewcollection#getItemOrNullObject_key_)|Obtém uma exibição de planilha usando seu nome.|
|[NotAvailableErrorCellValue](/javascript/api/excel/excel.notavailableerrorcellvalue)|[basicType](/javascript/api/excel/excel.notavailableerrorcellvalue#basicType)|Representa o valor que seria retornado por `Range.valueTypes` uma célula com esse valor.|
||[basicValue](/javascript/api/excel/excel.notavailableerrorcellvalue#basicValue)|Representa o valor que seria retornado por `Range.values` uma célula com esse valor.|
||[errorType](/javascript/api/excel/excel.notavailableerrorcellvalue#errorType)|Representa o tipo de `ErrorCellValue` .|
||[tipo](/javascript/api/excel/excel.notavailableerrorcellvalue#type)|Representa o tipo desse valor de célula.|
|[NullErrorCellValue](/javascript/api/excel/excel.nullerrorcellvalue)|[basicType](/javascript/api/excel/excel.nullerrorcellvalue#basicType)|Representa o valor que seria retornado por `Range.valueTypes` uma célula com esse valor.|
||[basicValue](/javascript/api/excel/excel.nullerrorcellvalue#basicValue)|Representa o valor que seria retornado por `Range.values` uma célula com esse valor.|
||[errorType](/javascript/api/excel/excel.nullerrorcellvalue#errorType)|Representa o tipo de `ErrorCellValue` .|
||[tipo](/javascript/api/excel/excel.nullerrorcellvalue#type)|Representa o tipo desse valor de célula.|
|[NumErrorCellValue](/javascript/api/excel/excel.numerrorcellvalue)|[basicType](/javascript/api/excel/excel.numerrorcellvalue#basicType)|Representa o valor que seria retornado por `Range.valueTypes` uma célula com esse valor.|
||[basicValue](/javascript/api/excel/excel.numerrorcellvalue#basicValue)|Representa o valor que seria retornado por `Range.values` uma célula com esse valor.|
||[errorType](/javascript/api/excel/excel.numerrorcellvalue#errorType)|Representa o tipo de `ErrorCellValue` .|
||[tipo](/javascript/api/excel/excel.numerrorcellvalue#type)|Representa o tipo desse valor de célula.|
|[PivotLayout](/javascript/api/excel/excel.pivotlayout)|[getCell(dataHierarchy: DataPivotHierarchy \| string, rowItems: Array<PivotItem \| string>, columnItems: Array<PivotItem \| string>)](/javascript/api/excel/excel.pivotlayout#getCell_dataHierarchy__rowItems__columnItems_)|Obtém uma célula exclusiva na tabela dinâmica com base em uma hierarquia de dados, bem como os itens de linha e coluna de suas respectivas hierarquias.|
||[pivotStyle](/javascript/api/excel/excel.pivotlayout#pivotStyle)|O estilo aplicado à Tabela Dinâmica.|
||[setStyle(style: string \| PivotTableStyle \| BuiltInPivotTableStyle)](/javascript/api/excel/excel.pivotlayout#setStyle_style_)|Define o estilo aplicado à Tabela Dinâmica.|
|[PivotTable](/javascript/api/excel/excel.pivottable)|[getDataSourceString()](/javascript/api/excel/excel.pivottable#getDataSourceString__)|Retorna a representação de cadeia de caracteres da fonte de dados da Tabela Dinâmica.|
||[getDataSourceType()](/javascript/api/excel/excel.pivottable#getDataSourceType__)|Obtém o tipo da fonte de dados da tabela dinâmica.|
|[PivotTableScopedCollection](/javascript/api/excel/excel.pivottablescopedcollection)|[getFirstOrNullObject()](/javascript/api/excel/excel.pivottablescopedcollection#getFirstOrNullObject__)|Obtém a primeira Tabela Dinâmica da coleção.|
|[Range](/javascript/api/excel/excel.range)|[getDependents()](/javascript/api/excel/excel.range#getDependents__)|Retorna um objeto que representa o intervalo que contém todos os dependentes de uma célula na mesma planilha ou `WorkbookRangeAreas` em várias planilhas.|
||[valuesAsJson](/javascript/api/excel/excel.range#valuesAsJson)|Uma representação JSON dos valores nas células nesse intervalo.|
|[RangeView](/javascript/api/excel/excel.rangeview)|[valuesAsJson](/javascript/api/excel/excel.rangeview#valuesAsJson)|Uma representação JSON dos valores nas células nesse intervalo.|
|[RefErrorCellValue](/javascript/api/excel/excel.referrorcellvalue)|[basicType](/javascript/api/excel/excel.referrorcellvalue#basicType)|Representa o valor que seria retornado por `Range.valueTypes` uma célula com esse valor.|
||[basicValue](/javascript/api/excel/excel.referrorcellvalue#basicValue)|Representa o valor que seria retornado por `Range.values` uma célula com esse valor.|
||[errorSubType](/javascript/api/excel/excel.referrorcellvalue#errorSubType)|Representa o tipo de `RefErrorCellValue` .|
||[errorType](/javascript/api/excel/excel.referrorcellvalue#errorType)|Representa o tipo de `ErrorCellValue` .|
||[tipo](/javascript/api/excel/excel.referrorcellvalue#type)|Representa o tipo desse valor de célula.|
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
|[Segmentação de dados](/javascript/api/excel/excel.slicer)|[nameInFormula](/javascript/api/excel/excel.slicer#nameInFormula)|Representa o nome da segmentação de dados usada na fórmula.|
||[setStyle(style: string \| SlicerStyle \| BuiltInSlicerStyle)](/javascript/api/excel/excel.slicer#setStyle_style_)|Define o estilo aplicado à slicer.|
||[slicerStyle](/javascript/api/excel/excel.slicer#slicerStyle)|O estilo aplicado à slicer.|
|[SpillErrorCellValue](/javascript/api/excel/excel.spillerrorcellvalue)|[basicType](/javascript/api/excel/excel.spillerrorcellvalue#basicType)|Representa o valor que seria retornado por `Range.valueTypes` uma célula com esse valor.|
||[basicValue](/javascript/api/excel/excel.spillerrorcellvalue#basicValue)|Representa o valor que seria retornado por `Range.values` uma célula com esse valor.|
||[errorSubType](/javascript/api/excel/excel.spillerrorcellvalue#errorSubType)|Representa o tipo de `SpillErrorCellValue` .|
||[errorType](/javascript/api/excel/excel.spillerrorcellvalue#errorType)|Representa o tipo de `ErrorCellValue` .|
||[spilledColumns](/javascript/api/excel/excel.spillerrorcellvalue#spilledColumns)|Representa o número de colunas que seriam derramadas se não houvesse #SPILL! .|
||[spilledRows](/javascript/api/excel/excel.spillerrorcellvalue#spilledRows)|Representa o número de linhas que seriam derramadas se não houvesse #SPILL! .|
||[tipo](/javascript/api/excel/excel.spillerrorcellvalue#type)|Representa o tipo desse valor de célula.|
|[StringCellValue](/javascript/api/excel/excel.stringcellvalue)|[basicType](/javascript/api/excel/excel.stringcellvalue#basicType)|Representa o valor que seria retornado por `Range.valueTypes` uma célula com esse valor.|
||[basicValue](/javascript/api/excel/excel.stringcellvalue#basicValue)|Representa o valor que seria retornado por `Range.values` uma célula com esse valor.|
||[tipo](/javascript/api/excel/excel.stringcellvalue#type)|Representa o tipo desse valor de célula.|
|[Table](/javascript/api/excel/excel.table)|[clearStyle()](/javascript/api/excel/excel.table#clearStyle__)|Altera a tabela para usar o estilo de tabela padrão.|
||[onFiltered](/javascript/api/excel/excel.table#onFiltered)|Ocorre quando um filtro é aplicado em uma tabela específica.|
||[setStyle(style: string \| TableStyle \| BuiltInTableStyle)](/javascript/api/excel/excel.table#setStyle_style_)|Define o estilo aplicado à tabela.|
||[tableStyle](/javascript/api/excel/excel.table#tableStyle)|O estilo aplicado à tabela.|
|[TableCollection](/javascript/api/excel/excel.tablecollection)|[onFiltered](/javascript/api/excel/excel.tablecollection#onFiltered)|Ocorre quando um filtro é aplicado em qualquer tabela em uma pasta de trabalho ou em uma planilha.|
|[TableColumn](/javascript/api/excel/excel.tablecolumn)|[valuesAsJson](/javascript/api/excel/excel.tablecolumn#valuesAsJson)|Uma representação JSON dos valores nas células nesta coluna de tabela.|
|[TableFilteredEventArgs](/javascript/api/excel/excel.tablefilteredeventargs)|[tableId](/javascript/api/excel/excel.tablefilteredeventargs#tableId)|Obtém a ID da tabela na qual o filtro é aplicado.|
||[tipo](/javascript/api/excel/excel.tablefilteredeventargs#type)|Obtém o tipo do evento.|
||[worksheetId](/javascript/api/excel/excel.tablefilteredeventargs#worksheetId)|Obtém a ID da planilha que contém a tabela.|
|[TableRow](/javascript/api/excel/excel.tablerow)|[valuesAsJson](/javascript/api/excel/excel.tablerow#valuesAsJson)|Uma representação JSON dos valores nas células nesta linha de tabela.|
|[ValueErrorCellValue](/javascript/api/excel/excel.valueerrorcellvalue)|[basicType](/javascript/api/excel/excel.valueerrorcellvalue#basicType)|Representa o valor que seria retornado por `Range.valueTypes` uma célula com esse valor.|
||[basicValue](/javascript/api/excel/excel.valueerrorcellvalue#basicValue)|Representa o valor que seria retornado por `Range.values` uma célula com esse valor.|
||[errorSubType](/javascript/api/excel/excel.valueerrorcellvalue#errorSubType)|Representa o tipo de `ValueErrorCellValue` .|
||[errorType](/javascript/api/excel/excel.valueerrorcellvalue#errorType)|Representa o tipo de `ErrorCellValue` .|
||[tipo](/javascript/api/excel/excel.valueerrorcellvalue#type)|Representa o tipo desse valor de célula.|
|[ValueTypeNotAvailableCellValue](/javascript/api/excel/excel.valuetypenotavailablecellvalue)|[basicType](/javascript/api/excel/excel.valuetypenotavailablecellvalue#basicType)|Representa o valor que seria retornado por `Range.valueTypes` uma célula com esse valor.|
||[basicValue](/javascript/api/excel/excel.valuetypenotavailablecellvalue#basicValue)|Representa o valor que seria retornado por `Range.values` uma célula com esse valor.|
||[tipo](/javascript/api/excel/excel.valuetypenotavailablecellvalue#type)|Representa o tipo desse valor de célula.|
|[WebImageCellValue](/javascript/api/excel/excel.webimagecellvalue)|[address](/javascript/api/excel/excel.webimagecellvalue#address)|Representa a URL da qual a imagem será baixada.|
||[altText](/javascript/api/excel/excel.webimagecellvalue#altText)|Representa o texto alternativo que pode ser usado em cenários de acessibilidade para descrever o que a imagem representa.|
||[attribution](/javascript/api/excel/excel.webimagecellvalue#attribution)|Representa informações de atribuição para descrever os requisitos de origem e licença para usar essa imagem.|
||[basicType](/javascript/api/excel/excel.webimagecellvalue#basicType)|Representa o valor que seria retornado por `Range.valueTypes` uma célula com esse valor.|
||[basicValue](/javascript/api/excel/excel.webimagecellvalue#basicValue)|Representa o valor que seria retornado por `Range.values` uma célula com esse valor.|
||[provider](/javascript/api/excel/excel.webimagecellvalue#provider)|Representa informações que descrevem a entidade ou indivíduo que forneceu a imagem.|
||[relatedImagesAddress](/javascript/api/excel/excel.webimagecellvalue#relatedImagesAddress)|Representa a URL de uma página da Web com imagens consideradas relacionadas a este `WebImageCellValue` .|
||[tipo](/javascript/api/excel/excel.webimagecellvalue#type)|Representa o tipo desse valor de célula.|
|[Workbook](/javascript/api/excel/excel.workbook)|[linkedDataTypes](/javascript/api/excel/excel.workbook#linkedDataTypes)|Retorna uma coleção de tipos de dados vinculados que fazem parte da lista de trabalho.|
||[showPivotFieldList](/javascript/api/excel/excel.workbook#showPivotFieldList)|Especifica se o painel de lista de campos da Tabela Dinâmica é mostrado no nível da lista de trabalho.|
||[tasks](/javascript/api/excel/excel.workbook#tasks)|Retorna uma coleção de tarefas que estão presentes na workbook.|
||[use1904DateSystem](/javascript/api/excel/excel.workbook#use1904DateSystem)|True se a pasta de trabalho usar o sistema de dados 1904.|
|[Worksheet](/javascript/api/excel/excel.worksheet)|[onFiltered](/javascript/api/excel/excel.worksheet#onFiltered)|Ocorre quando um filtro é aplicado em uma planilha específica.|
||[tasks](/javascript/api/excel/excel.worksheet#tasks)|Retorna uma coleção de tarefas presentes na planilha.|
|[WorksheetCollection](/javascript/api/excel/excel.worksheetcollection)|[addFromBase64(base64File: string, sheetNamesToInsert?: string[], positionType?: Excel.WorksheetPositionType, relativeTo?: Worksheet \| string)](/javascript/api/excel/excel.worksheetcollection#addFromBase64_base64File__sheetNamesToInsert__positionType__relativeTo_)|Insere as planilhas especificadas de uma pasta de trabalho na pasta de trabalho atual.|
||[onFiltered](/javascript/api/excel/excel.worksheetcollection#onFiltered)|Ocorre quando filtro de uma planilha é aplicado na pasta de trabalho.|
|[WorksheetFilteredEventArgs](/javascript/api/excel/excel.worksheetfilteredeventargs)|[tipo](/javascript/api/excel/excel.worksheetfilteredeventargs#type)|Obtém o tipo do evento.|
||[worksheetId](/javascript/api/excel/excel.worksheetfilteredeventargs#worksheetId)|Obtém a ID da planilha na qual o filtro é aplicado.|
|[WorksheetProtection](/javascript/api/excel/excel.worksheetprotection)|[allowEditRanges](/javascript/api/excel/excel.worksheetprotection#allowEditRanges)|Especifica o `AllowEditRangeCollection` encontrado nesta planilha.|
||[canPauseProtection](/javascript/api/excel/excel.worksheetprotection#canPauseProtection)|Especifica se a proteção pode ser pausada para esta planilha.|
||[checkPassword(password?: string)](/javascript/api/excel/excel.worksheetprotection#checkPassword_password_)|Especifica se a senha pode ser usada para desbloquear a proteção da planilha.|
||[isPasswordProtected](/javascript/api/excel/excel.worksheetprotection#isPasswordProtected)|Especifica se a planilha está protegida por senha.|
||[isPaused](/javascript/api/excel/excel.worksheetprotection#isPaused)|Especifica se a proteção da planilha está pausada.|
||[pauseProtection(password?: string)](/javascript/api/excel/excel.worksheetprotection#pauseProtection_password_)|Pausa a proteção da planilha para o objeto de planilha determinado para o usuário em uma determinada sessão.|
||[resumeProtection()](/javascript/api/excel/excel.worksheetprotection#resumeProtection__)|Retoma a proteção da planilha para o objeto de planilha determinado para o usuário em uma determinada sessão.|
||[setPassword(password?: string)](/javascript/api/excel/excel.worksheetprotection#setPassword_password_)|Altera a senha associada ao `WorksheetProtection` objeto.|
||[updateOptions(options: Excel. WorksheetProtectionOptions)](/javascript/api/excel/excel.worksheetprotection#updateOptions_options_)|Altere as opções de proteção da planilha associadas ao `WorksheetProtection` objeto.|
|[WorksheetProtectionChangedEventArgs](/javascript/api/excel/excel.worksheetprotectionchangedeventargs)|[allowEditRangesChanged](/javascript/api/excel/excel.worksheetprotectionchangedeventargs#allowEditRangesChanged)|Especifica se algum dos `AllowEditRange` objetos foi alterado.|
||[protectionOptionsChanged](/javascript/api/excel/excel.worksheetprotectionchangedeventargs#protectionOptionsChanged)|Especifica se o `WorksheetProtectionOptions` foi alterado.|
||[sheetPasswordChanged](/javascript/api/excel/excel.worksheetprotectionchangedeventargs#sheetPasswordChanged)|Especifica se a senha da planilha foi alterada.|

## <a name="see-also"></a>Confira também

- [Documentação deReferência da API JavaScript do Excel](/javascript/api/excel?view=excel-js-preview&preserve-view=true)
- [Conjuntos de requisitos da API JavaScript do Excel](excel-api-requirement-sets.md)
