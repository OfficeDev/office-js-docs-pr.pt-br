---
title: APIs de visualização do JavaScript para Excel
description: Detalhes sobre as próximas Excel APIs JavaScript.
ms.date: 12/08/2021
ms.prod: excel
ms.localizationpriority: medium
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
| Table styles | Fornece controle para fonte, borda, cor de preenchimento e outros aspectos dos estilos de tabela. | [Tabela](/javascript/api/excel/excel.table), [Tabela Dinâmica](/javascript/api/excel/excel.pivottable), [Slicer](/javascript/api/excel/excel.slicer) |
| Proteção de planilha | Impedir que usuários não autorizados mudem para intervalos especificados em uma planilha. | [WorksheetProtection](/javascript/api/excel/excel.worksheetprotection), [AllowEditRange](/javascript/api/excel/excel.alloweditrange), [AllowEditRangeCollection](/javascript/api/excel/excel.alloweditrangecollection), [AllowEditRangeOptions](/javascript/api/excel/excel.alloweditrangeoptions) |

## <a name="api-list"></a>Lista de API

A tabela a seguir lista as Excel APIs JavaScript atualmente em visualização. Para uma lista completa de todas as EXCEL JavaScript (incluindo APIs de visualização e APIs lançadas anteriormente), consulte todas as [APIs javascript Excel JavaScript](/javascript/api/excel?view=excel-js-preview&preserve-view=true).

| Classe | Campos | Descrição |
|:---|:---|:---|
|[AllowEditRange](/javascript/api/excel/excel.alloweditrange)|[address](/javascript/api/excel/excel.alloweditrange#excel-excel-alloweditrange-address-member)|Especifica o intervalo associado ao objeto.|
||[delete()](/javascript/api/excel/excel.alloweditrange#excel-excel-alloweditrange-delete-member(1))|Exclui esse objeto do `AllowEditRangeCollection`.|
||[isPasswordProtected](/javascript/api/excel/excel.alloweditrange#excel-excel-alloweditrange-ispasswordprotected-member)|Especifica se a senha `AllowEditRange` está protegida.|
||[pauseProtection(password?: string)](/javascript/api/excel/excel.alloweditrange#excel-excel-alloweditrange-pauseprotection-member(1))|Pausa a proteção da planilha para o objeto `AllowEditRange` dado para o usuário em uma determinada sessão.|
||[setPassword(password?: string)](/javascript/api/excel/excel.alloweditrange#excel-excel-alloweditrange-setpassword-member(1))|Altera a senha associada ao `AllowEditRange`.|
||[title](/javascript/api/excel/excel.alloweditrange#excel-excel-alloweditrange-title-member)|Especifica o título do objeto.|
|[AllowEditRangeCollection](/javascript/api/excel/excel.alloweditrangecollection)|[add(title: string, rangeAddress: string, options?: Excel. AllowEditRangeOptions)](/javascript/api/excel/excel.alloweditrangecollection#excel-excel-alloweditrangecollection-add-member(1))|Adiciona um `AllowEditRange` objeto à coleção.|
||[getCount()](/javascript/api/excel/excel.alloweditrangecollection#excel-excel-alloweditrangecollection-getcount-member(1))|Retorna o número de `AllowEditRange` objetos na coleção.|
||[getItem(key: string)](/javascript/api/excel/excel.alloweditrangecollection#excel-excel-alloweditrangecollection-getitem-member(1))|Obtém `AllowEditRange` o objeto pelo título.|
||[getItemAt(index: number)](/javascript/api/excel/excel.alloweditrangecollection#excel-excel-alloweditrangecollection-getitemat-member(1))|Retorna um `AllowEditRange` objeto pelo índice na coleção.|
||[getItemOrNullObject(key: string)](/javascript/api/excel/excel.alloweditrangecollection#excel-excel-alloweditrangecollection-getitemornullobject-member(1))|Obtém `AllowEditRange` o objeto pelo título.|
||[items](/javascript/api/excel/excel.alloweditrangecollection#excel-excel-alloweditrangecollection-items-member)|Obtém os itens filhos carregados nesta coleção.|
||[pauseProtection(password: string)](/javascript/api/excel/excel.alloweditrangecollection#excel-excel-alloweditrangecollection-pauseprotection-member(1))|Pausa a proteção da planilha para todos os `AllowEditRange` objetos da coleção que têm a senha dada para o usuário em uma determinada sessão.|
|[AllowEditRangeOptions](/javascript/api/excel/excel.alloweditrangeoptions)|[senha](/javascript/api/excel/excel.alloweditrangeoptions#excel-excel-alloweditrangeoptions-password-member)|A senha associada ao `AllowEditRange`.|
|[ArrayCellValue](/javascript/api/excel/excel.arraycellvalue)|[basicType](/javascript/api/excel/excel.arraycellvalue#excel-excel-arraycellvalue-basictype-member)|Representa o valor que seria retornado por uma `Range.valueTypes` célula com esse valor.|
||[basicValue](/javascript/api/excel/excel.arraycellvalue#excel-excel-arraycellvalue-basicvalue-member)|Representa o valor que seria retornado por uma `Range.values` célula com esse valor.|
||[elements](/javascript/api/excel/excel.arraycellvalue#excel-excel-arraycellvalue-elements-member)|Representa os elementos da matriz.|
||[tipo](/javascript/api/excel/excel.arraycellvalue#excel-excel-arraycellvalue-type-member)|Representa o tipo desse valor de célula.|
|[BlockedErrorCellValue](/javascript/api/excel/excel.blockederrorcellvalue)|[basicType](/javascript/api/excel/excel.blockederrorcellvalue#excel-excel-blockederrorcellvalue-basictype-member)|Representa o valor que seria retornado por uma `Range.valueTypes` célula com esse valor.|
||[basicValue](/javascript/api/excel/excel.blockederrorcellvalue#excel-excel-blockederrorcellvalue-basicvalue-member)|Representa o valor que seria retornado por uma `Range.values` célula com esse valor.|
||[errorSubType](/javascript/api/excel/excel.blockederrorcellvalue#excel-excel-blockederrorcellvalue-errorsubtype-member)|Representa o tipo de `BlockedErrorCellValue`.|
||[errorType](/javascript/api/excel/excel.blockederrorcellvalue#excel-excel-blockederrorcellvalue-errortype-member)|Representa o tipo de `ErrorCellValue`.|
||[tipo](/javascript/api/excel/excel.blockederrorcellvalue#excel-excel-blockederrorcellvalue-type-member)|Representa o tipo desse valor de célula.|
|[BooleanCellValue](/javascript/api/excel/excel.booleancellvalue)|[basicType](/javascript/api/excel/excel.booleancellvalue#excel-excel-booleancellvalue-basictype-member)|Representa o valor que seria retornado por uma `Range.valueTypes` célula com esse valor.|
||[basicValue](/javascript/api/excel/excel.booleancellvalue#excel-excel-booleancellvalue-basicvalue-member)|Representa o valor que seria retornado por uma `Range.values` célula com esse valor.|
||[tipo](/javascript/api/excel/excel.booleancellvalue#excel-excel-booleancellvalue-type-member)|Representa o tipo desse valor de célula.|
|[BusyErrorCellValue](/javascript/api/excel/excel.busyerrorcellvalue)|[basicType](/javascript/api/excel/excel.busyerrorcellvalue#excel-excel-busyerrorcellvalue-basictype-member)|Representa o valor que seria retornado por uma `Range.valueTypes` célula com esse valor.|
||[basicValue](/javascript/api/excel/excel.busyerrorcellvalue#excel-excel-busyerrorcellvalue-basicvalue-member)|Representa o valor que seria retornado por uma `Range.values` célula com esse valor.|
||[errorSubType](/javascript/api/excel/excel.busyerrorcellvalue#excel-excel-busyerrorcellvalue-errorsubtype-member)|Representa o tipo de `BusyErrorCellValue`.|
||[errorType](/javascript/api/excel/excel.busyerrorcellvalue#excel-excel-busyerrorcellvalue-errortype-member)|Representa o tipo de `ErrorCellValue`.|
||[tipo](/javascript/api/excel/excel.busyerrorcellvalue#excel-excel-busyerrorcellvalue-type-member)|Representa o tipo desse valor de célula.|
|[CalcErrorCellValue](/javascript/api/excel/excel.calcerrorcellvalue)|[basicType](/javascript/api/excel/excel.calcerrorcellvalue#excel-excel-calcerrorcellvalue-basictype-member)|Representa o valor que seria retornado por uma `Range.valueTypes` célula com esse valor.|
||[basicValue](/javascript/api/excel/excel.calcerrorcellvalue#excel-excel-calcerrorcellvalue-basicvalue-member)|Representa o valor que seria retornado por uma `Range.values` célula com esse valor.|
||[errorSubType](/javascript/api/excel/excel.calcerrorcellvalue#excel-excel-calcerrorcellvalue-errorsubtype-member)|Representa o tipo de `CalcErrorCellValue`.|
||[errorType](/javascript/api/excel/excel.calcerrorcellvalue#excel-excel-calcerrorcellvalue-errortype-member)|Representa o tipo de `ErrorCellValue`.|
||[tipo](/javascript/api/excel/excel.calcerrorcellvalue#excel-excel-calcerrorcellvalue-type-member)|Representa o tipo desse valor de célula.|
|[CardLayoutListSection](/javascript/api/excel/excel.cardlayoutlistsection)|[layout](/javascript/api/excel/excel.cardlayoutlistsection#excel-excel-cardlayoutlistsection-layout-member)|Representa o tipo de layout desta seção.|
|[CardLayoutPropertyReference](/javascript/api/excel/excel.cardlayoutpropertyreference)|[property](/javascript/api/excel/excel.cardlayoutpropertyreference#excel-excel-cardlayoutpropertyreference-property-member)|O nome da propriedade referenciada pelo layout do cartão.|
|[CardLayoutSectionStandardProperties](/javascript/api/excel/excel.cardlayoutsectionstandardproperties)|[collapsed](/javascript/api/excel/excel.cardlayoutsectionstandardproperties#excel-excel-cardlayoutsectionstandardproperties-collapsed-member)|Representa se esta seção do cartão está inicialmente recolhido.|
||[collapsible](/javascript/api/excel/excel.cardlayoutsectionstandardproperties#excel-excel-cardlayoutsectionstandardproperties-collapsible-member)|Representa se essa seção do cartão é retutível.|
||[properties](/javascript/api/excel/excel.cardlayoutsectionstandardproperties#excel-excel-cardlayoutsectionstandardproperties-properties-member)|Representa os nomes das propriedades nesta seção.|
||[title](/javascript/api/excel/excel.cardlayoutsectionstandardproperties#excel-excel-cardlayoutsectionstandardproperties-title-member)|Representa o título desta seção do cartão.|
|[CardLayoutStandardProperties](/javascript/api/excel/excel.cardlayoutstandardproperties)|[mainImage](/javascript/api/excel/excel.cardlayoutstandardproperties#excel-excel-cardlayoutstandardproperties-mainimage-member)|Especifica uma propriedade que será usada como a imagem principal do cartão.|
||[sections](/javascript/api/excel/excel.cardlayoutstandardproperties#excel-excel-cardlayoutstandardproperties-sections-member)|Representa as seções do cartão.|
||[subTitle](/javascript/api/excel/excel.cardlayoutstandardproperties#excel-excel-cardlayoutstandardproperties-subtitle-member)|Representa uma especificação de qual propriedade contém o subtítulo do cartão.|
||[title](/javascript/api/excel/excel.cardlayoutstandardproperties#excel-excel-cardlayoutstandardproperties-title-member)|Representa o título do cartão ou a especificação de qual propriedade contém o título do cartão.|
|[CardLayoutTableSection](/javascript/api/excel/excel.cardlayouttablesection)|[layout](/javascript/api/excel/excel.cardlayouttablesection#excel-excel-cardlayouttablesection-layout-member)|Representa o tipo de layout desta seção.|
|[CellValueAttributionAttributes](/javascript/api/excel/excel.cellvalueattributionattributes)|[licenseAddress](/javascript/api/excel/excel.cellvalueattributionattributes#excel-excel-cellvalueattributionattributes-licenseaddress-member)|Representa uma URL para uma licença ou fonte que descreve como essa propriedade pode ser usada.|
||[licenseText](/javascript/api/excel/excel.cellvalueattributionattributes#excel-excel-cellvalueattributionattributes-licensetext-member)|Representa um nome para a licença que rege essa propriedade.|
||[sourceAddress](/javascript/api/excel/excel.cellvalueattributionattributes#excel-excel-cellvalueattributionattributes-sourceaddress-member)|Representa uma URL para a origem do `CellValue`.|
||[sourceText](/javascript/api/excel/excel.cellvalueattributionattributes#excel-excel-cellvalueattributionattributes-sourcetext-member)|Representa um nome para a origem do `CellValue`.|
|[CellValuePropertyMetadata](/javascript/api/excel/excel.cellvaluepropertymetadata)|[attribution](/javascript/api/excel/excel.cellvaluepropertymetadata#excel-excel-cellvaluepropertymetadata-attribution-member)|Representa informações de atribuição para descrever os requisitos de origem e licença para usar essa propriedade.|
||[excludeFrom](/javascript/api/excel/excel.cellvaluepropertymetadata#excel-excel-cellvaluepropertymetadata-excludefrom-member)|Representa de quais recursos essa propriedade é excluída.|
||[sub-rótulo](/javascript/api/excel/excel.cellvaluepropertymetadata#excel-excel-cellvaluepropertymetadata-sublabel-member)|Representa o sub-rótulo dessa propriedade mostrado no exibição de cartão.|
|[CellValuePropertyMetadataExclusions](/javascript/api/excel/excel.cellvaluepropertymetadataexclusions)|[autoComplete](/javascript/api/excel/excel.cellvaluepropertymetadataexclusions#excel-excel-cellvaluepropertymetadataexclusions-autocomplete-member)|True representa que a propriedade é excluída das propriedades mostradas pela conclusão automática.|
||[calcCompare](/javascript/api/excel/excel.cellvaluepropertymetadataexclusions#excel-excel-cellvaluepropertymetadataexclusions-calccompare-member)|True representa que a propriedade é excluída das propriedades usadas para comparar valores de células durante o recalcal.|
||[cardView](/javascript/api/excel/excel.cellvaluepropertymetadataexclusions#excel-excel-cellvaluepropertymetadataexclusions-cardview-member)|True representa que a propriedade é excluída das propriedades mostradas pelo exibição de cartão.|
||[dotNotation](/javascript/api/excel/excel.cellvaluepropertymetadataexclusions#excel-excel-cellvaluepropertymetadataexclusions-dotnotation-member)|True representa que a propriedade é excluída das propriedades que podem ser acessadas por meio da função FIELDVALUE.|
|[CellValueProviderAttributes](/javascript/api/excel/excel.cellvalueproviderattributes)|[description](/javascript/api/excel/excel.cellvalueproviderattributes#excel-excel-cellvalueproviderattributes-description-member)|Representa a propriedade de descrição do provedor usada no exibição de cartão se nenhum logotipo for especificado.|
||[logoSourceAddress](/javascript/api/excel/excel.cellvalueproviderattributes#excel-excel-cellvalueproviderattributes-logosourceaddress-member)|Representa uma URL usada para baixar uma imagem que será usada como um logotipo no exibição de cartão.|
||[logoTargetAddress](/javascript/api/excel/excel.cellvalueproviderattributes#excel-excel-cellvalueproviderattributes-logotargetaddress-member)|Representa uma URL que será o destino de navegação se o usuário clicar no elemento logo no modo de exibição de cartão.|
|[Comentário](/javascript/api/excel/excel.comment)|[assignTask(assignee: Identity)](/javascript/api/excel/excel.comment#excel-excel-comment-assigntask-member(1))|Atribui a tarefa anexada ao comentário ao usuário dado como um destinatário.|
||[getTask()](/javascript/api/excel/excel.comment#excel-excel-comment-gettask-member(1))|Obtém a tarefa associada a este comentário.|
||[getTaskOrNullObject()](/javascript/api/excel/excel.comment#excel-excel-comment-gettaskornullobject-member(1))|Obtém a tarefa associada a este comentário.|
|[CommentReply](/javascript/api/excel/excel.commentreply)|[assignTask(assignee: Identity)](/javascript/api/excel/excel.commentreply#excel-excel-commentreply-assigntask-member(1))|Atribui a tarefa anexada ao comentário ao usuário determinado como o único destinatário.|
||[getTask()](/javascript/api/excel/excel.commentreply#excel-excel-commentreply-gettask-member(1))|Obtém a tarefa associada ao thread desta resposta de comentário.|
||[getTaskOrNullObject()](/javascript/api/excel/excel.commentreply#excel-excel-commentreply-gettaskornullobject-member(1))|Obtém a tarefa associada ao thread desta resposta de comentário.|
|[ConnectErrorCellValue](/javascript/api/excel/excel.connecterrorcellvalue)|[basicType](/javascript/api/excel/excel.connecterrorcellvalue#excel-excel-connecterrorcellvalue-basictype-member)|Representa o valor que seria retornado por uma `Range.valueTypes` célula com esse valor.|
||[basicValue](/javascript/api/excel/excel.connecterrorcellvalue#excel-excel-connecterrorcellvalue-basicvalue-member)|Representa o valor que seria retornado por uma `Range.values` célula com esse valor.|
||[errorSubType](/javascript/api/excel/excel.connecterrorcellvalue#excel-excel-connecterrorcellvalue-errorsubtype-member)|Representa o tipo de `ConnectErrorCellValue`.|
||[errorType](/javascript/api/excel/excel.connecterrorcellvalue#excel-excel-connecterrorcellvalue-errortype-member)|Representa o tipo de `ErrorCellValue`.|
||[tipo](/javascript/api/excel/excel.connecterrorcellvalue#excel-excel-connecterrorcellvalue-type-member)|Representa o tipo desse valor de célula.|
|[Div0ErrorCellValue](/javascript/api/excel/excel.div0errorcellvalue)|[basicType](/javascript/api/excel/excel.div0errorcellvalue#excel-excel-div0errorcellvalue-basictype-member)|Representa o valor que seria retornado por uma `Range.valueTypes` célula com esse valor.|
||[basicValue](/javascript/api/excel/excel.div0errorcellvalue#excel-excel-div0errorcellvalue-basicvalue-member)|Representa o valor que seria retornado por uma `Range.values` célula com esse valor.|
||[errorType](/javascript/api/excel/excel.div0errorcellvalue#excel-excel-div0errorcellvalue-errortype-member)|Representa o tipo de `ErrorCellValue`.|
||[tipo](/javascript/api/excel/excel.div0errorcellvalue#excel-excel-div0errorcellvalue-type-member)|Representa o tipo desse valor de célula.|
|[DocumentTask](/javascript/api/excel/excel.documenttask)|[assignees](/javascript/api/excel/excel.documenttask#excel-excel-documenttask-assignees-member)|Retorna uma coleção de atribuídos da tarefa.|
||[changes](/javascript/api/excel/excel.documenttask#excel-excel-documenttask-changes-member)|Obtém os registros de alteração da tarefa.|
||[comment](/javascript/api/excel/excel.documenttask#excel-excel-documenttask-comment-member)|Obtém o comentário associado à tarefa.|
||[completedBy](/javascript/api/excel/excel.documenttask#excel-excel-documenttask-completedby-member)|Obtém o usuário mais recente para ter concluído a tarefa.|
||[completedDateTime](/javascript/api/excel/excel.documenttask#excel-excel-documenttask-completeddatetime-member)|Obtém a data e a hora em que a tarefa foi concluída.|
||[createdBy](/javascript/api/excel/excel.documenttask#excel-excel-documenttask-createdby-member)|Obtém o usuário que criou a tarefa.|
||[createdDateTime](/javascript/api/excel/excel.documenttask#excel-excel-documenttask-createddatetime-member)|Obtém a data e a hora em que a tarefa foi criada.|
||[id](/javascript/api/excel/excel.documenttask#excel-excel-documenttask-id-member)|Obtém a ID da tarefa.|
||[percentComplete](/javascript/api/excel/excel.documenttask#excel-excel-documenttask-percentcomplete-member)|Especifica a porcentagem de conclusão da tarefa.|
||[priority](/javascript/api/excel/excel.documenttask#excel-excel-documenttask-priority-member)|Especifica a prioridade da tarefa.|
||[setStartAndDueDateTime(startDateTime: Date, dueDateTime: Date)](/javascript/api/excel/excel.documenttask#excel-excel-documenttask-setstartandduedatetime-member(1))|Altera o início e as datas de vencimento da tarefa.|
||[startAndDueDateTime](/javascript/api/excel/excel.documenttask#excel-excel-documenttask-startandduedatetime-member)|Obtém ou define a data e a hora em que a tarefa deve começar e deve ser final.|
||[title](/javascript/api/excel/excel.documenttask#excel-excel-documenttask-title-member)|Especifica o título da tarefa.|
|[DocumentTaskChange](/javascript/api/excel/excel.documenttaskchange)|[assignee](/javascript/api/excel/excel.documenttaskchange#excel-excel-documenttaskchange-assignee-member)|Representa o usuário atribuído à tarefa para `assign` um tipo de registro de alteração ou o usuário não atribuído da tarefa para `unassign` um tipo de registro de alteração.|
||[changedBy](/javascript/api/excel/excel.documenttaskchange#excel-excel-documenttaskchange-changedby-member)|Representa o usuário que criou ou alterou a tarefa.|
||[commentId](/javascript/api/excel/excel.documenttaskchange#excel-excel-documenttaskchange-commentid-member)|Representa a ID do `Comment` ou ao `CommentReply` qual a alteração da tarefa está ancorada.|
||[createdDateTime](/javascript/api/excel/excel.documenttaskchange#excel-excel-documenttaskchange-createddatetime-member)|Representa a data e a hora de criação do registro de alteração de tarefa.|
||[dueDateTime](/javascript/api/excel/excel.documenttaskchange#excel-excel-documenttaskchange-duedatetime-member)|Representa a data e a hora de vencimento da tarefa, no fuso horário UTC.|
||[id](/javascript/api/excel/excel.documenttaskchange#excel-excel-documenttaskchange-id-member)|ID do registro de alteração de tarefa.|
||[percentComplete](/javascript/api/excel/excel.documenttaskchange#excel-excel-documenttaskchange-percentcomplete-member)|Representa a porcentagem de conclusão da tarefa.|
||[priority](/javascript/api/excel/excel.documenttaskchange#excel-excel-documenttaskchange-priority-member)|Representa a prioridade da tarefa.|
||[startDateTime](/javascript/api/excel/excel.documenttaskchange#excel-excel-documenttaskchange-startdatetime-member)|Representa a data e a hora de início da tarefa, no fuso horário UTC.|
||[title](/javascript/api/excel/excel.documenttaskchange#excel-excel-documenttaskchange-title-member)|Representa o título da tarefa.|
||[tipo](/javascript/api/excel/excel.documenttaskchange#excel-excel-documenttaskchange-type-member)|Representa o tipo de ação do registro de alteração de tarefa.|
||[undoHistoryId](/javascript/api/excel/excel.documenttaskchange#excel-excel-documenttaskchange-undohistoryid-member)|Representa a `DocumentTaskChange.id` propriedade que foi desfeita para o tipo `undo` de registro de alteração.|
|[DocumentTaskChangeCollection](/javascript/api/excel/excel.documenttaskchangecollection)|[getCount()](/javascript/api/excel/excel.documenttaskchangecollection#excel-excel-documenttaskchangecollection-getcount-member(1))|Obtém o número de registros de alteração na coleção da tarefa.|
||[getItemAt(index: number)](/javascript/api/excel/excel.documenttaskchangecollection#excel-excel-documenttaskchangecollection-getitemat-member(1))|Obtém um registro de alteração de tarefa usando seu índice na coleção.|
||[items](/javascript/api/excel/excel.documenttaskchangecollection#excel-excel-documenttaskchangecollection-items-member)|Obtém os itens filhos carregados nesta coleção.|
|[DocumentTaskCollection](/javascript/api/excel/excel.documenttaskcollection)|[getCount()](/javascript/api/excel/excel.documenttaskcollection#excel-excel-documenttaskcollection-getcount-member(1))|Obtém o número de tarefas na coleção.|
||[getItem(key: string)](/javascript/api/excel/excel.documenttaskcollection#excel-excel-documenttaskcollection-getitem-member(1))|Obtém uma tarefa usando sua ID.|
||[getItemAt(index: number)](/javascript/api/excel/excel.documenttaskcollection#excel-excel-documenttaskcollection-getitemat-member(1))|Obtém uma tarefa pelo índice na coleção.|
||[getItemOrNullObject(key: string)](/javascript/api/excel/excel.documenttaskcollection#excel-excel-documenttaskcollection-getitemornullobject-member(1))|Obtém uma tarefa usando sua ID.|
||[items](/javascript/api/excel/excel.documenttaskcollection#excel-excel-documenttaskcollection-items-member)|Obtém os itens filhos carregados nesta coleção.|
|[DocumentTaskSchedule](/javascript/api/excel/excel.documenttaskschedule)|[dueDateTime](/javascript/api/excel/excel.documenttaskschedule#excel-excel-documenttaskschedule-duedatetime-member)|Obtém a data e a hora de vencimento da tarefa.|
||[startDateTime](/javascript/api/excel/excel.documenttaskschedule#excel-excel-documenttaskschedule-startdatetime-member)|Obtém a data e a hora em que a tarefa deve começar.|
|[DoubleCellValue](/javascript/api/excel/excel.doublecellvalue)|[basicType](/javascript/api/excel/excel.doublecellvalue#excel-excel-doublecellvalue-basictype-member)|Representa o valor que seria retornado por uma `Range.valueTypes` célula com esse valor.|
||[basicValue](/javascript/api/excel/excel.doublecellvalue#excel-excel-doublecellvalue-basicvalue-member)|Representa o valor que seria retornado por uma `Range.values` célula com esse valor.|
||[tipo](/javascript/api/excel/excel.doublecellvalue#excel-excel-doublecellvalue-type-member)|Representa o tipo desse valor de célula.|
|[EmptyCellValue](/javascript/api/excel/excel.emptycellvalue)|[basicType](/javascript/api/excel/excel.emptycellvalue#excel-excel-emptycellvalue-basictype-member)|Representa o valor que seria retornado por uma `Range.valueTypes` célula com esse valor.|
||[basicValue](/javascript/api/excel/excel.emptycellvalue#excel-excel-emptycellvalue-basicvalue-member)|Representa o valor que seria retornado por uma `Range.values` célula com esse valor.|
||[tipo](/javascript/api/excel/excel.emptycellvalue#excel-excel-emptycellvalue-type-member)|Representa o tipo desse valor de célula.|
|[EntityCardLayout](/javascript/api/excel/excel.entitycardlayout)|[layout](/javascript/api/excel/excel.entitycardlayout#excel-excel-entitycardlayout-layout-member)|Representa o tipo desse layout.|
|[EntityCellValue](/javascript/api/excel/excel.entitycellvalue)|[basicType](/javascript/api/excel/excel.entitycellvalue#excel-excel-entitycellvalue-basictype-member)|Representa o valor que seria retornado por uma `Range.valueTypes` célula com esse valor.|
||[basicValue](/javascript/api/excel/excel.entitycellvalue#excel-excel-entitycellvalue-basicvalue-member)|Representa o valor que seria retornado por uma `Range.values` célula com esse valor.|
||[cardLayout](/javascript/api/excel/excel.entitycellvalue#excel-excel-entitycellvalue-cardlayout-member)|Representa o layout dessa entidade no exibição de cartão.|
||[properties: { [key: string]](/javascript/api/excel/excel.entitycellvalue#excel-excel-entitycellvalue-properties-member)|Representa as propriedades dessa entidade e seus metadados.|
||[text](/javascript/api/excel/excel.entitycellvalue#excel-excel-entitycellvalue-text-member)|Representa o texto mostrado quando uma célula com esse valor é renderizada.|
||[tipo](/javascript/api/excel/excel.entitycellvalue#excel-excel-entitycellvalue-type-member)|Representa o tipo desse valor de célula.|
|[FieldErrorCellValue](/javascript/api/excel/excel.fielderrorcellvalue)|[basicType](/javascript/api/excel/excel.fielderrorcellvalue#excel-excel-fielderrorcellvalue-basictype-member)|Representa o valor que seria retornado por uma `Range.valueTypes` célula com esse valor.|
||[basicValue](/javascript/api/excel/excel.fielderrorcellvalue#excel-excel-fielderrorcellvalue-basicvalue-member)|Representa o valor que seria retornado por uma `Range.values` célula com esse valor.|
||[errorSubType](/javascript/api/excel/excel.fielderrorcellvalue#excel-excel-fielderrorcellvalue-errorsubtype-member)|Representa o tipo de `FieldErrorCellValue`.|
||[errorType](/javascript/api/excel/excel.fielderrorcellvalue#excel-excel-fielderrorcellvalue-errortype-member)|Representa o tipo de `ErrorCellValue`.|
||[tipo](/javascript/api/excel/excel.fielderrorcellvalue#excel-excel-fielderrorcellvalue-type-member)|Representa o tipo desse valor de célula.|
|[FormattedNumberCellValue](/javascript/api/excel/excel.formattednumbercellvalue)|[basicType](/javascript/api/excel/excel.formattednumbercellvalue#excel-excel-formattednumbercellvalue-basictype-member)|Representa o valor que seria retornado por uma `Range.valueTypes` célula com esse valor.|
||[basicValue](/javascript/api/excel/excel.formattednumbercellvalue#excel-excel-formattednumbercellvalue-basicvalue-member)|Representa o valor que seria retornado por uma `Range.values` célula com esse valor.|
||[numberFormat](/javascript/api/excel/excel.formattednumbercellvalue#excel-excel-formattednumbercellvalue-numberformat-member)|Retorna a cadeia de caracteres de formato de número usada para exibir esse valor.|
||[tipo](/javascript/api/excel/excel.formattednumbercellvalue#excel-excel-formattednumbercellvalue-type-member)|Representa o tipo desse valor de célula.|
|[GettingDataErrorCellValue](/javascript/api/excel/excel.gettingdataerrorcellvalue)|[basicType](/javascript/api/excel/excel.gettingdataerrorcellvalue#excel-excel-gettingdataerrorcellvalue-basictype-member)|Representa o valor que seria retornado por uma `Range.valueTypes` célula com esse valor.|
||[basicValue](/javascript/api/excel/excel.gettingdataerrorcellvalue#excel-excel-gettingdataerrorcellvalue-basicvalue-member)|Representa o valor que seria retornado por uma `Range.values` célula com esse valor.|
||[errorType](/javascript/api/excel/excel.gettingdataerrorcellvalue#excel-excel-gettingdataerrorcellvalue-errortype-member)|Representa o tipo de `ErrorCellValue`.|
||[tipo](/javascript/api/excel/excel.gettingdataerrorcellvalue#excel-excel-gettingdataerrorcellvalue-type-member)|Representa o tipo desse valor de célula.|
|[Identidade](/javascript/api/excel/excel.identity)|[displayName](/javascript/api/excel/excel.identity#excel-excel-identity-displayname-member)|Representa o nome para exibição do usuário.|
||[email](/javascript/api/excel/excel.identity#excel-excel-identity-email-member)|Representa o endereço de email do usuário.|
||[id](/javascript/api/excel/excel.identity#excel-excel-identity-id-member)|Representa a ID exclusiva do usuário.|
|[IdentityCollection](/javascript/api/excel/excel.identitycollection)|[add(assignee: Identity)](/javascript/api/excel/excel.identitycollection#excel-excel-identitycollection-add-member(1))|Adiciona uma identidade de usuário à coleção.|
||[clear()](/javascript/api/excel/excel.identitycollection#excel-excel-identitycollection-clear-member(1))|Remove todas as identidades de usuário da coleção.|
||[getCount()](/javascript/api/excel/excel.identitycollection#excel-excel-identitycollection-getcount-member(1))|Obtém o número de itens na coleção.|
||[getItemAt(index: number)](/javascript/api/excel/excel.identitycollection#excel-excel-identitycollection-getitemat-member(1))|Obtém uma identidade de usuário de documento usando seu índice na coleção.|
||[items](/javascript/api/excel/excel.identitycollection#excel-excel-identitycollection-items-member)|Obtém os itens filhos carregados nesta coleção.|
||[remove(assignee: Identity)](/javascript/api/excel/excel.identitycollection#excel-excel-identitycollection-remove-member(1))|Remove uma identidade de usuário da coleção.|
|[IdentityEntity](/javascript/api/excel/excel.identityentity)|[displayName](/javascript/api/excel/excel.identityentity#excel-excel-identityentity-displayname-member)|Representa o nome para exibição do usuário.|
||[email](/javascript/api/excel/excel.identityentity#excel-excel-identityentity-email-member)|Representa o endereço de email do usuário.|
||[id](/javascript/api/excel/excel.identityentity#excel-excel-identityentity-id-member)|Representa a ID exclusiva do usuário.|
|[LinkedDataType](/javascript/api/excel/excel.linkeddatatype)|[dataProvider](/javascript/api/excel/excel.linkeddatatype#excel-excel-linkeddatatype-dataprovider-member)|O nome do provedor de dados do tipo de dados vinculado.|
||[lastRefreshed](/javascript/api/excel/excel.linkeddatatype#excel-excel-linkeddatatype-lastrefreshed-member)|A data e a hora do fuso horário local desde que a lista de trabalho foi aberta quando o tipo de dados vinculado foi atualizado pela última vez.|
||[name](/javascript/api/excel/excel.linkeddatatype#excel-excel-linkeddatatype-name-member)|O nome do tipo de dados vinculado.|
||[periodicRefreshInterval](/javascript/api/excel/excel.linkeddatatype#excel-excel-linkeddatatype-periodicrefreshinterval-member)|A frequência, em segundos, na qual o tipo de dados vinculado é atualizado se `refreshMode` estiver definido como "Periódico".|
||[refreshMode](/javascript/api/excel/excel.linkeddatatype#excel-excel-linkeddatatype-refreshmode-member)|O mecanismo pelo qual os dados do tipo de dados vinculados são recuperados.|
||[requestRefresh()](/javascript/api/excel/excel.linkeddatatype#excel-excel-linkeddatatype-requestrefresh-member(1))|Faz uma solicitação para atualizar o tipo de dados vinculado.|
||[requestSetRefreshMode(refreshMode: Excel. LinkedDataTypeRefreshMode)](/javascript/api/excel/excel.linkeddatatype#excel-excel-linkeddatatype-requestsetrefreshmode-member(1))|Faz uma solicitação para alterar o modo de atualização para esse tipo de dados vinculado.|
||[serviceId](/javascript/api/excel/excel.linkeddatatype#excel-excel-linkeddatatype-serviceid-member)|A ID exclusiva do tipo de dados vinculado.|
||[supportedRefreshModes](/javascript/api/excel/excel.linkeddatatype#excel-excel-linkeddatatype-supportedrefreshmodes-member)|Retorna uma matriz com todos os modos de atualização suportados pelo tipo de dados vinculado.|
|[LinkedDataTypeAddedEventArgs](/javascript/api/excel/excel.linkeddatatypeaddedeventargs)|[serviceId](/javascript/api/excel/excel.linkeddatatypeaddedeventargs#excel-excel-linkeddatatypeaddedeventargs-serviceid-member)|A ID exclusiva do novo tipo de dados vinculado.|
||[source](/javascript/api/excel/excel.linkeddatatypeaddedeventargs#excel-excel-linkeddatatypeaddedeventargs-source-member)|Obtém a origem do evento.|
||[tipo](/javascript/api/excel/excel.linkeddatatypeaddedeventargs#excel-excel-linkeddatatypeaddedeventargs-type-member)|Obtém o tipo do evento.|
|[LinkedDataTypeCollection](/javascript/api/excel/excel.linkeddatatypecollection)|[getCount()](/javascript/api/excel/excel.linkeddatatypecollection#excel-excel-linkeddatatypecollection-getcount-member(1))|Obtém o número de tipos de dados vinculados na coleção.|
||[getItem(key: number)](/javascript/api/excel/excel.linkeddatatypecollection#excel-excel-linkeddatatypecollection-getitem-member(1))|Obtém um tipo de dados vinculado por ID de serviço.|
||[getItemAt(index: number)](/javascript/api/excel/excel.linkeddatatypecollection#excel-excel-linkeddatatypecollection-getitemat-member(1))|Obtém um tipo de dados vinculado pelo índice na coleção.|
||[getItemOrNullObject(key: number)](/javascript/api/excel/excel.linkeddatatypecollection#excel-excel-linkeddatatypecollection-getitemornullobject-member(1))|Obtém um tipo de dados vinculado por ID.|
||[items](/javascript/api/excel/excel.linkeddatatypecollection#excel-excel-linkeddatatypecollection-items-member)|Obtém os itens filhos carregados nesta coleção.|
||[requestRefreshAll()](/javascript/api/excel/excel.linkeddatatypecollection#excel-excel-linkeddatatypecollection-requestrefreshall-member(1))|Faz uma solicitação para atualizar todos os tipos de dados vinculados na coleção.|
|[LinkedEntityCellValue](/javascript/api/excel/excel.linkedentitycellvalue)|[basicType](/javascript/api/excel/excel.linkedentitycellvalue#excel-excel-linkedentitycellvalue-basictype-member)|Representa o valor que seria retornado por uma `Range.valueTypes` célula com esse valor.|
||[basicValue](/javascript/api/excel/excel.linkedentitycellvalue#excel-excel-linkedentitycellvalue-basicvalue-member)|Representa o valor que seria retornado por uma `Range.values` célula com esse valor.|
||[id](/javascript/api/excel/excel.linkedentitycellvalue#excel-excel-linkedentitycellvalue-id-member)|Representa a fonte de serviço que forneceu as informações nesse valor.|
||[properties: { [key: string]: CellValue & { propertyMetadata](/javascript/api/excel/excel.linkedentitycellvalue#excel-excel-linkedentitycellvalue-properties-member)|Representa as propriedades dessa entidade e seus metadados.|
||[propertyMetadata](/javascript/api/excel/excel.linkedentitycellvalue#excel-excel-linkedentitycellvalue-propertymetadata-member)||
||[provider](/javascript/api/excel/excel.linkedentitycellvalue#excel-excel-linkedentitycellvalue-provider-member)|Representa informações que descrevem o serviço que forneceu a imagem.|
||[texto](/javascript/api/excel/excel.linkedentitycellvalue#excel-excel-linkedentitycellvalue-text-member)|Representa o texto mostrado quando uma célula com esse valor é renderizada.|
||[tipo](/javascript/api/excel/excel.linkedentitycellvalue#excel-excel-linkedentitycellvalue-type-member)|Representa o tipo desse valor de célula.|
|[LinkedEntityId](/javascript/api/excel/excel.linkedentityid)|[culture](/javascript/api/excel/excel.linkedentityid#excel-excel-linkedentityid-culture-member)|Representa qual cultura de idioma foi usada para criar isso `CellValue`.|
||[domainId](/javascript/api/excel/excel.linkedentityid#excel-excel-linkedentityid-domainid-member)|Representa um domínio específico de um serviço usado para criar o `CellValue`.|
||[entityId](/javascript/api/excel/excel.linkedentityid#excel-excel-linkedentityid-entityid-member)|Representa um identificador específico de um serviço usado para criar o `CellValue`.|
||[serviceId](/javascript/api/excel/excel.linkedentityid#excel-excel-linkedentityid-serviceid-member)|Representa qual serviço foi usado para criar o `CellValue`.|
|[NameErrorCellValue](/javascript/api/excel/excel.nameerrorcellvalue)|[basicType](/javascript/api/excel/excel.nameerrorcellvalue#excel-excel-nameerrorcellvalue-basictype-member)|Representa o valor que seria retornado por uma `Range.valueTypes` célula com esse valor.|
||[basicValue](/javascript/api/excel/excel.nameerrorcellvalue#excel-excel-nameerrorcellvalue-basicvalue-member)|Representa o valor que seria retornado por uma `Range.values` célula com esse valor.|
||[errorType](/javascript/api/excel/excel.nameerrorcellvalue#excel-excel-nameerrorcellvalue-errortype-member)|Representa o tipo de `ErrorCellValue`.|
||[tipo](/javascript/api/excel/excel.nameerrorcellvalue#excel-excel-nameerrorcellvalue-type-member)|Representa o tipo desse valor de célula.|
|[NamedItem](/javascript/api/excel/excel.nameditem)|[valueAsJson](/javascript/api/excel/excel.nameditem#excel-excel-nameditem-valueasjson-member)|Uma representação JSON dos valores neste item nomeado.|
|[NamedItemArrayValues](/javascript/api/excel/excel.nameditemarrayvalues)|[valuesAsJson](/javascript/api/excel/excel.nameditemarrayvalues#excel-excel-nameditemarrayvalues-valuesasjson-member)|Uma representação JSON dos valores nas células nesse intervalo.|
|[NamedSheetViewCollection](/javascript/api/excel/excel.namedsheetviewcollection)|[getItemOrNullObject(key: string)](/javascript/api/excel/excel.namedsheetviewcollection#excel-excel-namedsheetviewcollection-getitemornullobject-member(1))|Obtém uma exibição de planilha usando seu nome.|
|[NotAvailableErrorCellValue](/javascript/api/excel/excel.notavailableerrorcellvalue)|[basicType](/javascript/api/excel/excel.notavailableerrorcellvalue#excel-excel-notavailableerrorcellvalue-basictype-member)|Representa o valor que seria retornado por uma `Range.valueTypes` célula com esse valor.|
||[basicValue](/javascript/api/excel/excel.notavailableerrorcellvalue#excel-excel-notavailableerrorcellvalue-basicvalue-member)|Representa o valor que seria retornado por uma `Range.values` célula com esse valor.|
||[errorType](/javascript/api/excel/excel.notavailableerrorcellvalue#excel-excel-notavailableerrorcellvalue-errortype-member)|Representa o tipo de `ErrorCellValue`.|
||[tipo](/javascript/api/excel/excel.notavailableerrorcellvalue#excel-excel-notavailableerrorcellvalue-type-member)|Representa o tipo desse valor de célula.|
|[NullErrorCellValue](/javascript/api/excel/excel.nullerrorcellvalue)|[basicType](/javascript/api/excel/excel.nullerrorcellvalue#excel-excel-nullerrorcellvalue-basictype-member)|Representa o valor que seria retornado por uma `Range.valueTypes` célula com esse valor.|
||[basicValue](/javascript/api/excel/excel.nullerrorcellvalue#excel-excel-nullerrorcellvalue-basicvalue-member)|Representa o valor que seria retornado por uma `Range.values` célula com esse valor.|
||[errorType](/javascript/api/excel/excel.nullerrorcellvalue#excel-excel-nullerrorcellvalue-errortype-member)|Representa o tipo de `ErrorCellValue`.|
||[tipo](/javascript/api/excel/excel.nullerrorcellvalue#excel-excel-nullerrorcellvalue-type-member)|Representa o tipo desse valor de célula.|
|[NumErrorCellValue](/javascript/api/excel/excel.numerrorcellvalue)|[basicType](/javascript/api/excel/excel.numerrorcellvalue#excel-excel-numerrorcellvalue-basictype-member)|Representa o valor que seria retornado por uma `Range.valueTypes` célula com esse valor.|
||[basicValue](/javascript/api/excel/excel.numerrorcellvalue#excel-excel-numerrorcellvalue-basicvalue-member)|Representa o valor que seria retornado por uma `Range.values` célula com esse valor.|
||[errorType](/javascript/api/excel/excel.numerrorcellvalue#excel-excel-numerrorcellvalue-errortype-member)|Representa o tipo de `ErrorCellValue`.|
||[tipo](/javascript/api/excel/excel.numerrorcellvalue#excel-excel-numerrorcellvalue-type-member)|Representa o tipo desse valor de célula.|
|[PivotLayout](/javascript/api/excel/excel.pivotlayout)|[getCell(dataHierarchy: DataPivotHierarchy \| string, rowItems: Array<PivotItem \| string>, columnItems: Array<PivotItem \| string>)](/javascript/api/excel/excel.pivotlayout#excel-excel-pivotlayout-getcell-member(1))|Obtém uma célula exclusiva na tabela dinâmica com base em uma hierarquia de dados, bem como os itens de linha e coluna de suas respectivas hierarquias.|
||[pivotStyle](/javascript/api/excel/excel.pivotlayout#excel-excel-pivotlayout-pivotstyle-member)|O estilo aplicado à Tabela Dinâmica.|
||[setStyle(style: string \| PivotTableStyle \| BuiltInPivotTableStyle)](/javascript/api/excel/excel.pivotlayout#excel-excel-pivotlayout-setstyle-member(1))|Define o estilo aplicado à Tabela Dinâmica.|
|[PivotTable](/javascript/api/excel/excel.pivottable)|[getDataSourceString()](/javascript/api/excel/excel.pivottable#excel-excel-pivottable-getdatasourcestring-member(1))|Retorna a representação de cadeia de caracteres da fonte de dados da Tabela Dinâmica.|
||[getDataSourceType()](/javascript/api/excel/excel.pivottable#excel-excel-pivottable-getdatasourcetype-member(1))|Obtém o tipo da fonte de dados da tabela dinâmica.|
|[PivotTableScopedCollection](/javascript/api/excel/excel.pivottablescopedcollection)|[getFirstOrNullObject()](/javascript/api/excel/excel.pivottablescopedcollection#excel-excel-pivottablescopedcollection-getfirstornullobject-member(1))|Obtém a primeira Tabela Dinâmica da coleção.|
|[PlaceholderErrorCellValue](/javascript/api/excel/excel.placeholdererrorcellvalue)|[basicType](/javascript/api/excel/excel.placeholdererrorcellvalue#excel-excel-placeholdererrorcellvalue-basictype-member)|Representa o valor que seria retornado por uma `Range.valueTypes` célula com esse valor.|
||[basicValue](/javascript/api/excel/excel.placeholdererrorcellvalue#excel-excel-placeholdererrorcellvalue-basicvalue-member)|Representa o valor que seria retornado por uma `Range.values` célula com esse valor.|
||[errorType](/javascript/api/excel/excel.placeholdererrorcellvalue#excel-excel-placeholdererrorcellvalue-errortype-member)|Representa o tipo de `ErrorCellValue`.|
||[target](/javascript/api/excel/excel.placeholdererrorcellvalue#excel-excel-placeholdererrorcellvalue-target-member)|`PlaceholderErrorCellValue` é usado durante o processamento, enquanto os dados são baixados.|
||[tipo](/javascript/api/excel/excel.placeholdererrorcellvalue#excel-excel-placeholdererrorcellvalue-type-member)|Representa o tipo desse valor de célula.|
|[Range](/javascript/api/excel/excel.range)|[getDependents()](/javascript/api/excel/excel.range#excel-excel-range-getdependents-member(1))|Retorna um `WorkbookRangeAreas` objeto que representa o intervalo que contém todos os dependentes de uma célula na mesma planilha ou em várias planilhas.|
||[valuesAsJson](/javascript/api/excel/excel.range#excel-excel-range-valuesasjson-member)|Uma representação JSON dos valores nas células nesse intervalo.|
|[RangeView](/javascript/api/excel/excel.rangeview)|[valuesAsJson](/javascript/api/excel/excel.rangeview#excel-excel-rangeview-valuesasjson-member)|Uma representação JSON dos valores nas células nesse intervalo.|
|[RefErrorCellValue](/javascript/api/excel/excel.referrorcellvalue)|[basicType](/javascript/api/excel/excel.referrorcellvalue#excel-excel-referrorcellvalue-basictype-member)|Representa o valor que seria retornado por uma `Range.valueTypes` célula com esse valor.|
||[basicValue](/javascript/api/excel/excel.referrorcellvalue#excel-excel-referrorcellvalue-basicvalue-member)|Representa o valor que seria retornado por uma `Range.values` célula com esse valor.|
||[errorSubType](/javascript/api/excel/excel.referrorcellvalue#excel-excel-referrorcellvalue-errorsubtype-member)|Representa o tipo de `RefErrorCellValue`.|
||[errorType](/javascript/api/excel/excel.referrorcellvalue#excel-excel-referrorcellvalue-errortype-member)|Representa o tipo de `ErrorCellValue`.|
||[tipo](/javascript/api/excel/excel.referrorcellvalue#excel-excel-referrorcellvalue-type-member)|Representa o tipo desse valor de célula.|
|[RefreshModeChangedEventArgs](/javascript/api/excel/excel.refreshmodechangedeventargs)|[refreshMode](/javascript/api/excel/excel.refreshmodechangedeventargs#excel-excel-refreshmodechangedeventargs-refreshmode-member)|O modo de atualização do tipo de dados vinculado.|
||[serviceId](/javascript/api/excel/excel.refreshmodechangedeventargs#excel-excel-refreshmodechangedeventargs-serviceid-member)|A ID exclusiva do objeto cujo modo de atualização foi alterado.|
||[source](/javascript/api/excel/excel.refreshmodechangedeventargs#excel-excel-refreshmodechangedeventargs-source-member)|Obtém a origem do evento.|
||[tipo](/javascript/api/excel/excel.refreshmodechangedeventargs#excel-excel-refreshmodechangedeventargs-type-member)|Obtém o tipo do evento.|
|[RefreshRequestCompletedEventArgs](/javascript/api/excel/excel.refreshrequestcompletedeventargs)|[atualizado](/javascript/api/excel/excel.refreshrequestcompletedeventargs#excel-excel-refreshrequestcompletedeventargs-refreshed-member)|Indica se a solicitação de atualização foi bem-sucedida.|
||[serviceId](/javascript/api/excel/excel.refreshrequestcompletedeventargs#excel-excel-refreshrequestcompletedeventargs-serviceid-member)|A ID exclusiva do objeto cuja solicitação de atualização foi concluída.|
||[source](/javascript/api/excel/excel.refreshrequestcompletedeventargs#excel-excel-refreshrequestcompletedeventargs-source-member)|Obtém a origem do evento.|
||[tipo](/javascript/api/excel/excel.refreshrequestcompletedeventargs#excel-excel-refreshrequestcompletedeventargs-type-member)|Obtém o tipo do evento.|
||[avisos](/javascript/api/excel/excel.refreshrequestcompletedeventargs#excel-excel-refreshrequestcompletedeventargs-warnings-member)|Uma matriz que contém quaisquer avisos gerados a partir da solicitação de atualização.|
|[Shape](/javascript/api/excel/excel.shape)|[displayName](/javascript/api/excel/excel.shape#excel-excel-shape-displayname-member)|Obtém o nome de exibição da forma.|
|[ShapeCollection](/javascript/api/excel/excel.shapecollection)|[addSvg(xml: string)](/javascript/api/excel/excel.shapecollection#excel-excel-shapecollection-addsvg-member(1))|Cria um gráfico vetorial escalável (SVG) de uma cadeia de caracteres XML e a adiciona à planilha.|
|[Segmentação de dados](/javascript/api/excel/excel.slicer)|[nameInFormula](/javascript/api/excel/excel.slicer#excel-excel-slicer-nameinformula-member)|Representa o nome da segmentação de dados usada na fórmula.|
||[setStyle(style: string \| SlicerStyle \| BuiltInSlicerStyle)](/javascript/api/excel/excel.slicer#excel-excel-slicer-setstyle-member(1))|Define o estilo aplicado à slicer.|
||[slicerStyle](/javascript/api/excel/excel.slicer#excel-excel-slicer-slicerstyle-member)|O estilo aplicado à slicer.|
|[SpillErrorCellValue](/javascript/api/excel/excel.spillerrorcellvalue)|[basicType](/javascript/api/excel/excel.spillerrorcellvalue#excel-excel-spillerrorcellvalue-basictype-member)|Representa o valor que seria retornado por uma `Range.valueTypes` célula com esse valor.|
||[basicValue](/javascript/api/excel/excel.spillerrorcellvalue#excel-excel-spillerrorcellvalue-basicvalue-member)|Representa o valor que seria retornado por uma `Range.values` célula com esse valor.|
||[errorSubType](/javascript/api/excel/excel.spillerrorcellvalue#excel-excel-spillerrorcellvalue-errorsubtype-member)|Representa o tipo de `SpillErrorCellValue`.|
||[errorType](/javascript/api/excel/excel.spillerrorcellvalue#excel-excel-spillerrorcellvalue-errortype-member)|Representa o tipo de `ErrorCellValue`.|
||[spilledColumns](/javascript/api/excel/excel.spillerrorcellvalue#excel-excel-spillerrorcellvalue-spilledcolumns-member)|Representa o número de colunas que seriam derramadas se não houvesse #SPILL! .|
||[spilledRows](/javascript/api/excel/excel.spillerrorcellvalue#excel-excel-spillerrorcellvalue-spilledrows-member)|Representa o número de linhas que seriam derramadas se não houvesse #SPILL! .|
||[tipo](/javascript/api/excel/excel.spillerrorcellvalue#excel-excel-spillerrorcellvalue-type-member)|Representa o tipo desse valor de célula.|
|[StringCellValue](/javascript/api/excel/excel.stringcellvalue)|[basicType](/javascript/api/excel/excel.stringcellvalue#excel-excel-stringcellvalue-basictype-member)|Representa o valor que seria retornado por uma `Range.valueTypes` célula com esse valor.|
||[basicValue](/javascript/api/excel/excel.stringcellvalue#excel-excel-stringcellvalue-basicvalue-member)|Representa o valor que seria retornado por uma `Range.values` célula com esse valor.|
||[tipo](/javascript/api/excel/excel.stringcellvalue#excel-excel-stringcellvalue-type-member)|Representa o tipo desse valor de célula.|
|[Table](/javascript/api/excel/excel.table)|[clearStyle()](/javascript/api/excel/excel.table#excel-excel-table-clearstyle-member(1))|Altera a tabela para usar o estilo de tabela padrão.|
||[onFiltered](/javascript/api/excel/excel.table#excel-excel-table-onfiltered-member)|Ocorre quando um filtro é aplicado em uma tabela específica.|
||[setStyle(style: string \| TableStyle \| BuiltInTableStyle)](/javascript/api/excel/excel.table#excel-excel-table-setstyle-member(1))|Define o estilo aplicado à tabela.|
||[tableStyle](/javascript/api/excel/excel.table#excel-excel-table-tablestyle-member)|O estilo aplicado à tabela.|
|[TableCollection](/javascript/api/excel/excel.tablecollection)|[onFiltered](/javascript/api/excel/excel.tablecollection#excel-excel-tablecollection-onfiltered-member)|Ocorre quando um filtro é aplicado em qualquer tabela em uma pasta de trabalho ou em uma planilha.|
|[TableColumn](/javascript/api/excel/excel.tablecolumn)|[valuesAsJson](/javascript/api/excel/excel.tablecolumn#excel-excel-tablecolumn-valuesasjson-member)|Uma representação JSON dos valores nas células nesta coluna de tabela.|
|[TableFilteredEventArgs](/javascript/api/excel/excel.tablefilteredeventargs)|[tableId](/javascript/api/excel/excel.tablefilteredeventargs#excel-excel-tablefilteredeventargs-tableid-member)|Obtém a ID da tabela na qual o filtro é aplicado.|
||[tipo](/javascript/api/excel/excel.tablefilteredeventargs#excel-excel-tablefilteredeventargs-type-member)|Obtém o tipo do evento.|
||[worksheetId](/javascript/api/excel/excel.tablefilteredeventargs#excel-excel-tablefilteredeventargs-worksheetid-member)|Obtém a ID da planilha que contém a tabela.|
|[TableRow](/javascript/api/excel/excel.tablerow)|[valuesAsJson](/javascript/api/excel/excel.tablerow#excel-excel-tablerow-valuesasjson-member)|Uma representação JSON dos valores nas células nesta linha de tabela.|
|[ValueErrorCellValue](/javascript/api/excel/excel.valueerrorcellvalue)|[basicType](/javascript/api/excel/excel.valueerrorcellvalue#excel-excel-valueerrorcellvalue-basictype-member)|Representa o valor que seria retornado por uma `Range.valueTypes` célula com esse valor.|
||[basicValue](/javascript/api/excel/excel.valueerrorcellvalue#excel-excel-valueerrorcellvalue-basicvalue-member)|Representa o valor que seria retornado por uma `Range.values` célula com esse valor.|
||[errorSubType](/javascript/api/excel/excel.valueerrorcellvalue#excel-excel-valueerrorcellvalue-errorsubtype-member)|Representa o tipo de `ValueErrorCellValue`.|
||[errorType](/javascript/api/excel/excel.valueerrorcellvalue#excel-excel-valueerrorcellvalue-errortype-member)|Representa o tipo de `ErrorCellValue`.|
||[tipo](/javascript/api/excel/excel.valueerrorcellvalue#excel-excel-valueerrorcellvalue-type-member)|Representa o tipo desse valor de célula.|
|[ValueTypeNotAvailableCellValue](/javascript/api/excel/excel.valuetypenotavailablecellvalue)|[basicType](/javascript/api/excel/excel.valuetypenotavailablecellvalue#excel-excel-valuetypenotavailablecellvalue-basictype-member)|Representa o valor que seria retornado por uma `Range.valueTypes` célula com esse valor.|
||[basicValue](/javascript/api/excel/excel.valuetypenotavailablecellvalue#excel-excel-valuetypenotavailablecellvalue-basicvalue-member)|Representa o valor que seria retornado por uma `Range.values` célula com esse valor.|
||[tipo](/javascript/api/excel/excel.valuetypenotavailablecellvalue#excel-excel-valuetypenotavailablecellvalue-type-member)|Representa o tipo desse valor de célula.|
|[WebImageCellValue](/javascript/api/excel/excel.webimagecellvalue)|[address](/javascript/api/excel/excel.webimagecellvalue#excel-excel-webimagecellvalue-address-member)|Representa a URL da qual a imagem será baixada.|
||[altText](/javascript/api/excel/excel.webimagecellvalue#excel-excel-webimagecellvalue-alttext-member)|Representa o texto alternativo que pode ser usado em cenários de acessibilidade para descrever o que a imagem representa.|
||[attribution](/javascript/api/excel/excel.webimagecellvalue#excel-excel-webimagecellvalue-attribution-member)|Representa informações de atribuição para descrever os requisitos de origem e licença para usar essa imagem.|
||[basicType](/javascript/api/excel/excel.webimagecellvalue#excel-excel-webimagecellvalue-basictype-member)|Representa o valor que seria retornado por uma `Range.valueTypes` célula com esse valor.|
||[basicValue](/javascript/api/excel/excel.webimagecellvalue#excel-excel-webimagecellvalue-basicvalue-member)|Representa o valor que seria retornado por uma `Range.values` célula com esse valor.|
||[provider](/javascript/api/excel/excel.webimagecellvalue#excel-excel-webimagecellvalue-provider-member)|Representa informações que descrevem a entidade ou indivíduo que forneceu a imagem.|
||[relatedImagesAddress](/javascript/api/excel/excel.webimagecellvalue#excel-excel-webimagecellvalue-relatedimagesaddress-member)|Representa a URL de uma página da Web com imagens consideradas relacionadas a este `WebImageCellValue`.|
||[tipo](/javascript/api/excel/excel.webimagecellvalue#excel-excel-webimagecellvalue-type-member)|Representa o tipo desse valor de célula.|
|[Workbook](/javascript/api/excel/excel.workbook)|[getLinkedEntityCellValue(linkedEntityCellValueId: LinkedEntityId)](/javascript/api/excel/excel.workbook#excel-excel-workbook-getlinkedentitycellvalue-member(1))|Retorna um `LinkedEntityCellValue` com base no `LinkedEntityId`fornecido .|
||[linkedDataTypes](/javascript/api/excel/excel.workbook#excel-excel-workbook-linkeddatatypes-member)|Retorna uma coleção de tipos de dados vinculados que fazem parte da lista de trabalho.|
||[showPivotFieldList](/javascript/api/excel/excel.workbook#excel-excel-workbook-showpivotfieldlist-member)|Especifica se o painel de lista de campos da Tabela Dinâmica é mostrado no nível da lista de trabalho.|
||[tasks](/javascript/api/excel/excel.workbook#excel-excel-workbook-tasks-member)|Retorna uma coleção de tarefas que estão presentes na workbook.|
||[use1904DateSystem](/javascript/api/excel/excel.workbook#excel-excel-workbook-use1904datesystem-member)|True se a pasta de trabalho usar o sistema de dados 1904.|
|[Worksheet](/javascript/api/excel/excel.worksheet)|[onFiltered](/javascript/api/excel/excel.worksheet#excel-excel-worksheet-onfiltered-member)|Ocorre quando um filtro é aplicado em uma planilha específica.|
||[tasks](/javascript/api/excel/excel.worksheet#excel-excel-worksheet-tasks-member)|Retorna uma coleção de tarefas presentes na planilha.|
|[WorksheetCollection](/javascript/api/excel/excel.worksheetcollection)|[addFromBase64(base64File: string, sheetNamesToInsert?: string[], positionType?: Excel.WorksheetPositionType, relativeTo?: Worksheet \| string)](/javascript/api/excel/excel.worksheetcollection#excel-excel-worksheetcollection-addfrombase64-member(1))|Insere as planilhas especificadas de uma pasta de trabalho na pasta de trabalho atual.|
||[onFiltered](/javascript/api/excel/excel.worksheetcollection#excel-excel-worksheetcollection-onfiltered-member)|Ocorre quando filtro de uma planilha é aplicado na pasta de trabalho.|
|[WorksheetFilteredEventArgs](/javascript/api/excel/excel.worksheetfilteredeventargs)|[tipo](/javascript/api/excel/excel.worksheetfilteredeventargs#excel-excel-worksheetfilteredeventargs-type-member)|Obtém o tipo do evento.|
||[worksheetId](/javascript/api/excel/excel.worksheetfilteredeventargs#excel-excel-worksheetfilteredeventargs-worksheetid-member)|Obtém a ID da planilha na qual o filtro é aplicado.|
|[WorksheetProtection](/javascript/api/excel/excel.worksheetprotection)|[allowEditRanges](/javascript/api/excel/excel.worksheetprotection#excel-excel-worksheetprotection-alloweditranges-member)|Especifica o objeto `AllowEditRangeCollection` encontrado nesta planilha.|
||[canPauseProtection](/javascript/api/excel/excel.worksheetprotection#excel-excel-worksheetprotection-canpauseprotection-member)|Especifica se a proteção pode ser pausada para esta planilha.|
||[checkPassword(password?: string)](/javascript/api/excel/excel.worksheetprotection#excel-excel-worksheetprotection-checkpassword-member(1))|Especifica se a senha pode ser usada para desbloquear a proteção da planilha.|
||[isPasswordProtected](/javascript/api/excel/excel.worksheetprotection#excel-excel-worksheetprotection-ispasswordprotected-member)|Especifica se a planilha está protegida por senha.|
||[isPaused](/javascript/api/excel/excel.worksheetprotection#excel-excel-worksheetprotection-ispaused-member)|Especifica se a proteção da planilha está pausada.|
||[pauseProtection(password?: string)](/javascript/api/excel/excel.worksheetprotection#excel-excel-worksheetprotection-pauseprotection-member(1))|Pausa a proteção da planilha para o objeto de planilha determinado para o usuário em uma determinada sessão.|
||[resumeProtection()](/javascript/api/excel/excel.worksheetprotection#excel-excel-worksheetprotection-resumeprotection-member(1))|Retoma a proteção da planilha para o objeto de planilha determinado para o usuário em uma determinada sessão.|
||[setPassword(password?: string)](/javascript/api/excel/excel.worksheetprotection#excel-excel-worksheetprotection-setpassword-member(1))|Altera a senha associada ao `WorksheetProtection` objeto.|
||[updateOptions(options: Excel. WorksheetProtectionOptions)](/javascript/api/excel/excel.worksheetprotection#excel-excel-worksheetprotection-updateoptions-member(1))|Altere as opções de proteção da planilha associadas ao `WorksheetProtection` objeto.|
|[WorksheetProtectionChangedEventArgs](/javascript/api/excel/excel.worksheetprotectionchangedeventargs)|[allowEditRangesChanged](/javascript/api/excel/excel.worksheetprotectionchangedeventargs#excel-excel-worksheetprotectionchangedeventargs-alloweditrangeschanged-member)|Especifica se algum dos objetos `AllowEditRange` foi alterado.|
||[protectionOptionsChanged](/javascript/api/excel/excel.worksheetprotectionchangedeventargs#excel-excel-worksheetprotectionchangedeventargs-protectionoptionschanged-member)|Especifica se o `WorksheetProtectionOptions` foi alterado.|
||[sheetPasswordChanged](/javascript/api/excel/excel.worksheetprotectionchangedeventargs#excel-excel-worksheetprotectionchangedeventargs-sheetpasswordchanged-member)|Especifica se a senha da planilha foi alterada.|

## <a name="see-also"></a>Confira também

- [Documentação deReferência da API JavaScript do Excel](/javascript/api/excel?view=excel-js-preview&preserve-view=true)
- [Conjuntos de requisitos da API JavaScript do Excel](excel-api-requirement-sets.md)
