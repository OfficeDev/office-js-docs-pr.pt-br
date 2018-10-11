
# <a name="item"></a>item

### <a name="officeofficemdcontextofficecontextmdmailboxofficecontextmailboxmditem"></a>[Office](office.md)[.context](office.context.md)[.mailbox](office.context.mailbox.md).item

O namespace `item` é usado para acessar a mensagem, a solicitação de reunião ou o compromisso selecionado no momento. Você pode determinar o tipo do `item` usando a propriedade [itemType](#itemtype-officemailboxenumsitemtypejavascriptapioutlook15officemailboxenumsitemtype).

##### <a name="requirements"></a>Requisitos

|Requisito| Valor|
|---|---|
|[Versão mínima do conjunto de requisitos de caixa de correio](/javascript/office/requirement-sets/outlook-api-requirement-sets)| 1.0|
|[Nível de permissão mínimo](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| Restrito|
|[Modo aplicável do Outlook](https://docs.microsoft.com/outlook/add-ins/#extension-points)| Redigir ou ler|

##### <a name="members-and-methods"></a>Membros e métodos

| Membro | Tipo |
|--------|------|
| [attachments](#attachments-arrayattachmentdetailsjavascriptapioutlook15officeattachmentdetails) | Membro |
| [bcc](#bcc-recipientsjavascriptapioutlook15officerecipients) | Membro |
| [body](#body-bodyjavascriptapioutlook15officebody) | Membro |
| [cc](#cc-arrayemailaddressdetailsjavascriptapioutlook15officeemailaddressdetailsrecipientsjavascriptapioutlook15officerecipients) | Membro |
| [conversationId](#nullable-conversationid-string) | Membro |
| [dateTimeCreated](#datetimecreated-date) | Membro |
| [dateTimeModified](#datetimemodified-date) | Membro |
| [end](#end-datetimejavascriptapioutlook15officetime) | Membro |
| [from](#from-emailaddressdetailsjavascriptapioutlook15officeemailaddressdetails) | Membro |
| [internetMessageId](#internetmessageid-string) | Membro |
| [itemClass](#itemclass-string) | Membro |
| [itemId](#nullable-itemid-string) | Membro |
| [itemType](#itemtype-officemailboxenumsitemtypejavascriptapioutlook15officemailboxenumsitemtype) | Membro |
| [location](#location-stringlocationjavascriptapioutlook15officelocation) | Membro |
| [normalizedSubject](#normalizedsubject-string) | Membro |
| [NotificationMessages](#notificationmessages-notificationmessagesjavascriptapioutlook15officenotificationmessages) | Membro |
| [optionalAttendees](#optionalattendees-arrayemailaddressdetailsjavascriptapioutlook15officeemailaddressdetailsrecipientsjavascriptapioutlook15officerecipients) | Membro |
| [organizer](#organizer-emailaddressdetailsjavascriptapioutlook15officeemailaddressdetails) | Membro |
| [requiredAttendees](#requiredattendees-arrayemailaddressdetailsjavascriptapioutlook15officeemailaddressdetailsrecipientsjavascriptapioutlook15officerecipients) | Membro |
| [sender](#sender-emailaddressdetailsjavascriptapioutlook15officeemailaddressdetails) | Membro |
| [start](#start-datetimejavascriptapioutlook15officetime) | Membro |
| [subject](#subject-stringsubjectjavascriptapioutlook15officesubject) | Membro |
| [to](#to-arrayemailaddressdetailsjavascriptapioutlook15officeemailaddressdetailsrecipientsjavascriptapioutlook15officerecipients) | Membro |
| [addFileAttachmentAsync](#addfileattachmentasyncuri-attachmentname-options-callback) | Método |
| [addItemAttachmentAsync](#additemattachmentasyncitemid-attachmentname-options-callback) | Método |
| [close](#close) | Método |
| [displayReplyAllForm](#displayreplyallformformdata) | Método |
| [displayReplyForm](#displayreplyformformdata) | Método |
| [getEntities](#getentities--entitiesjavascriptapioutlook15officeentities) | Método |
| [getEntitiesByType](#getentitiesbytypeentitytype--nullable-arraystringcontactjavascriptapioutlook15officecontactmeetingsuggestionjavascriptapioutlook15officemeetingsuggestionphonenumberjavascriptapioutlook15officephonenumbertasksuggestionjavascriptapioutlook15officetasksuggestion) | Método |
| [getFilteredEntitiesByName](#getfilteredentitiesbynamename--nullable-arraystringcontactjavascriptapioutlook15officecontactmeetingsuggestionjavascriptapioutlook15officemeetingsuggestionphonenumberjavascriptapioutlook15officephonenumbertasksuggestionjavascriptapioutlook15officetasksuggestion) | Método |
| [getRegExMatches](#getregexmatches--object) | Método |
| [getRegExMatchesByName](#getregexmatchesbynamename--nullable-array-string-) | Método |
| [getSelectedDataAsync](#getselecteddataasynccoerciontype-options-callback--string) | Método |
| [loadCustomPropertiesAsync](#loadcustompropertiesasynccallback-usercontext) | Método |
| [removeAttachmentAsync](#removeattachmentasyncattachmentid-options-callback) | Método |
| [saveAsync](#saveasyncoptions-callback) | Método |
| [setSelectedDataAsync](#setselecteddataasyncdata-options-callback) | Método |

### <a name="example"></a>Exemplo

O exemplo de código JavaScript a seguir mostra como acessar a propriedade `subject`  do item atual no Outlook.

```
// The initialize function is required for all apps.
Office.initialize = function () {
    // Checks for the DOM to load using the jQuery ready function.
    $(document).ready(function () {
    // After the DOM is loaded, app-specific code can run.
    var item = Office.context.mailbox.item;
    var subject = item.subject;
    // Continue with processing the subject of the current item,
    // which can be a message or appointment.
    });
}
```

### <a name="members"></a>Membros

#### <a name="attachments-arrayattachmentdetailsjavascriptapioutlook15officeattachmentdetails"></a>attachments :Array.<[AttachmentDetails](/javascript/api/outlook_1_5/office.attachmentdetails)>

Obtém uma matriz de anexos para o item. Somente modo de leitura.

> [!NOTE]
> Certos tipos de arquivos são bloqueados pelo Outlook devido a problemas potenciais de segurança e, portanto, não são retornados. Para obter mais informações, confira [Anexos bloqueados no Outlook](https://support.office.com/article/Blocked-attachments-in-Outlook-434752E1-02D3-4E90-9124-8B81E49A8519).

##### <a name="type"></a>Tipo:

*   Array.<[AttachmentDetails](/javascript/api/outlook_1_5/office.attachmentdetails)>

##### <a name="requirements"></a>Requisitos

|Requisito| Valor|
|---|---|
|[Versão mínima do conjunto de requisitos de caixa de correio](/javascript/office/requirement-sets/outlook-api-requirement-sets)| 1.0|
|[Nível de permissão mínimo](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| ReadItem|
|[Modo aplicável do Outlook](https://docs.microsoft.com/outlook/add-ins/#extension-points)| Leitura|

##### <a name="example"></a>Exemplo

O código a seguir cria uma sequência de caracteres HTML com detalhes de todos os anexos no item atual.

```
var _Item = Office.context.mailbox.item;
var outputString = "";

if (_Item.attachments.length > 0) {
  for (i = 0 ; i < _Item.attachments.length ; i++) {
    var _att = _Item.attachments[i];
    outputString += "<BR>" + i + ". Name: ";
    outputString += _att.name;
    outputString += "<BR>ID: " + _att.id;
    outputString += "<BR>contentType: " + _att.contentType;
    outputString += "<BR>size: " + _att.size;
    outputString += "<BR>attachmentType: " + _att.attachmentType;
    outputString += "<BR>isInline: " + _att.isInline;
  }
}

// Do something with outputString
```

####  <a name="bcc-recipientsjavascriptapioutlook15officerecipients"></a>bcc :[Recipients](/javascript/api/outlook_1_5/office.recipients)

Obtém um objeto que fornece os métodos para obter ou atualizar os destinatários na linha Cco (com cópia oculta) de uma mensagem. Somente modo de redação.

##### <a name="type"></a>Tipo:

*   [Destinatários](/javascript/api/outlook_1_5/office.recipients)

##### <a name="requirements"></a>Requisitos

|Requisito| Valor|
|---|---|
|[Versão mínima do conjunto de requisitos de caixa de correio](/javascript/office/requirement-sets/outlook-api-requirement-sets)| 1.1|
|[Nível de permissão mínimo](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| ReadItem|
|[Modo aplicável do Outlook](https://docs.microsoft.com/outlook/add-ins/#extension-points)| Redigir|

##### <a name="example"></a>Exemplo

```
Office.context.mailbox.item.bcc.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.bcc.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.bcc.getAsync(callback);

function callback(asyncResult) {
  var arrayOfBccRecipients = asyncResult.value;
}
```

####  <a name="body-bodyjavascriptapioutlook15officebody"></a>body :[Body](/javascript/api/outlook_1_5/office.body)

Obtém um objeto que fornece métodos para manipular o corpo de um item.

##### <a name="type"></a>Tipo:

*   [Corpo](/javascript/api/outlook_1_5/office.body)

##### <a name="requirements"></a>Requisitos

|Requisito| Valor|
|---|---|
|[Versão mínima do conjunto de requisitos de caixa de correio](/javascript/office/requirement-sets/outlook-api-requirement-sets)| 1.1|
|[Nível de permissão mínimo](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| ReadItem|
|[Modo aplicável do Outlook](https://docs.microsoft.com/outlook/add-ins/#extension-points)| Redigir ou ler|

####  <a name="cc-arrayemailaddressdetailsjavascriptapioutlook15officeemailaddressdetailsrecipientsjavascriptapioutlook15officerecipients"></a>cc :Array.<[EmailAddressDetails](/javascript/api/outlook_1_5/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook_1_5/office.recipients)

Fornece acesso aos destinatários Cc (com cópia) de uma mensagem. O tipo de objeto e o nível de acesso dependem do modo do item atual.

##### <a name="read-mode"></a>Modo de leitura

A propriedade `cc` retorna uma matriz que contém um objeto `EmailAddressDetails` para cada destinatário listado na linha **Cc** da mensagem. A coleção está limitada a um máximo de 100 membros.

##### <a name="compose-mode"></a>Modo de redação

A propriedade `cc` retorna um objeto `Recipients` que fornece métodos para obter ou atualizar os destinatários na linha **Cc** da mensagem.

##### <a name="type"></a>Tipo:

*   Array.<[EmailAddressDetails](/javascript/api/outlook_1_5/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook_1_5/office.recipients)

##### <a name="requirements"></a>Requisitos

|Requisito| Valor|
|---|---|
|[Versão mínima do conjunto de requisitos de caixa de correio](/javascript/office/requirement-sets/outlook-api-requirement-sets)| 1.0|
|[Nível de permissão mínimo](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| ReadItem|
|[Modo aplicável do Outlook](https://docs.microsoft.com/outlook/add-ins/#extension-points)| Redigir ou ler|

##### <a name="example"></a>Exemplo

```
Office.context.mailbox.item.cc.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.cc.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.cc.getAsync(callback);

function callback(asyncResult) {
  var arrayOfCcRecipients = asyncResult.value;
}
```

####  <a name="nullable-conversationid-string"></a>(nullable) conversationId :String

Obtém um identificador da conversa do e-mail que contém uma mensagem específica.

Você pode obter um número inteiro para esta propriedade se o aplicativo de e-mail estiver ativado nos formulários de leitura ou respostas nos formulários de redação. Se posteriormente o usuário alterar o assunto da mensagem de resposta, ao enviar a resposta, a ID da conversa daquela mensagem será alterada e o valor obtido anteriormente não mais se aplicará.

Você obterá nulo para esta propriedade para um novo item em um formulário de redação. Se o usuário definir um assunto e salvar o item, a propriedade `conversationId` retornará um valor.

##### <a name="type"></a>Tipo:

*   String

##### <a name="requirements"></a>Requisitos

|Requisito| Valor|
|---|---|
|[Versão mínima do conjunto de requisitos de caixa de correio](/javascript/office/requirement-sets/outlook-api-requirement-sets)| 1.0|
|[Nível de permissão mínimo](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| ReadItem|
|[Modo aplicável do Outlook](https://docs.microsoft.com/outlook/add-ins/#extension-points)| Redigir ou ler|

#### <a name="datetimecreated-date"></a>dateTimeCreated :Date

Obtém a data e a hora em que um item foi criado. Somente modo de leitura.

##### <a name="type"></a>Tipo:

*   Data

##### <a name="requirements"></a>Requisitos

|Requisito| Valor|
|---|---|
|[Versão mínima do conjunto de requisitos de caixa de correio](/javascript/office/requirement-sets/outlook-api-requirement-sets)| 1.0|
|[Nível de permissão mínimo](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| ReadItem|
|[Modo aplicável do Outlook](https://docs.microsoft.com/outlook/add-ins/#extension-points)| Leitura|

##### <a name="example"></a>Exemplo

```
var created = Office.context.mailbox.item.dateTimeCreated;
```

#### <a name="datetimemodified-date"></a>dateTimeModified :Date

Obtém a data e a hora em que um item foi alterado pela última vez. Somente modo de leitura.

> [!NOTE]
> Esse membro não é compatível com o Outlook para iOS ou o Outlook para Android.

##### <a name="type"></a>Tipo:

*   Date

##### <a name="requirements"></a>Requisitos

|Requisito| Valor|
|---|---|
|[Versão mínima do conjunto de requisitos de caixa de correio](/javascript/office/requirement-sets/outlook-api-requirement-sets)| 1.0|
|[Nível de permissão mínimo](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| ReadItem|
|[Modo aplicável do Outlook](https://docs.microsoft.com/outlook/add-ins/#extension-points)| Leitura|

##### <a name="example"></a>Exemplo

```
var modified = Office.context.mailbox.item.dateTimeModified;
```

####  <a name="end-datetimejavascriptapioutlook15officetime"></a>end :Date|[Time](/javascript/api/outlook_1_5/office.time)

Obtém ou define a data e a hora em que o compromisso deve terminar.

A propriedade `end` é expressa como um valor de data e hora no Tempo Universal Coordenado (UTC). Você pode usar o método [`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttimejavascriptapioutlook15officelocalclienttime) para converter o valor da propriedade de término para a data e a hora local do cliente.

##### <a name="read-mode"></a>Modo de leitura

A propriedade `end` retorna um objeto `Date`.

##### <a name="compose-mode"></a>Modo de redação

A propriedade `end` retorna um objeto `Time`.

Ao usar o método [`Time.setAsync`](/javascript/api/outlook_1_5/office.time#setasync-datetime--options--callback-) para definir a hora de término, deve-se usar o método [`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) para converter a hora local no cliente para UTC no servidor.

##### <a name="type"></a>Tipo:

*   Date | [Time](/javascript/api/outlook_1_5/office.time)

##### <a name="requirements"></a>Requisitos

|Requisito| Valor|
|---|---|
|[Versão mínima do conjunto de requisitos de caixa de correio](/javascript/office/requirement-sets/outlook-api-requirement-sets)| 1.0|
|[Nível de permissão mínimo](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| ReadItem|
|[Modo aplicável do Outlook](https://docs.microsoft.com/outlook/add-ins/#extension-points)| Redigir ou ler|

##### <a name="example"></a>Exemplo

O exemplo a seguir define a hora de término de um compromisso no modo de redação usando o método [`setAsync`](/javascript/api/outlook_1_5/office.time#setasync-datetime--options--callback-) do objeto `Time`.

```
var endTime = new Date("3/14/2015");
var options = {
  // Pass information that can be used
  // in the callback
     asyncContext: {verb:"Set"}
}
Office.context.mailbox.item.end.setAsync(endTime, options, function(result) {
  if (result.error) {
    console.debug(result.error);
  } else {
    // Access the asyncContext that was passed to the setAsync function
    console.debug("End Time " + result.asyncContext.verb);
  }
});
```

#### <a name="from-emailaddressdetailsjavascriptapioutlook15officeemailaddressdetails"></a>from :[EmailAddressDetails](/javascript/api/outlook_1_5/office.emailaddressdetails)

Obtém o endereço de e-mail do remetente de uma mensagem. Somente modo de leitura.

As propriedades `from` e [`sender`](#sender-emailaddressdetailsjavascriptapioutlook15officeemailaddressdetails) correspondem a mesma pessoa, a menos que a mensagem seja enviada por um representante. Nesse caso, a propriedade `from` corresponde ao representado e a propriedade sender ao representante.

> [!NOTE]
> A propriedade `recipientType` do objeto `EmailAddressDetails` na propriedade `from` é `undefined`.

##### <a name="type"></a>Tipo:

*   [EmailAddressDetails](/javascript/api/outlook_1_5/office.emailaddressdetails)

##### <a name="requirements"></a>Requisitos

|Requisito| Valor|
|---|---|
|[Versão mínima do conjunto de requisitos de caixa de correio](/javascript/office/requirement-sets/outlook-api-requirement-sets)| 1.0|
|[Nível de permissão mínimo](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| ReadItem|
|[Modo aplicável do Outlook](https://docs.microsoft.com/outlook/add-ins/#extension-points)| Leitura|

#### <a name="internetmessageid-string"></a>internetMessageId :String

Obtém o identificador de mensagem de Internet para uma mensagem de e-mail. Somente modo de leitura.

##### <a name="type"></a>Tipo:

*   String

##### <a name="requirements"></a>Requisitos

|Requisito| Valor|
|---|---|
|[Versão mínima do conjunto de requisitos de caixa de correio](/javascript/office/requirement-sets/outlook-api-requirement-sets)| 1.0|
|[Nível de permissão mínimo](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| ReadItem|
|[Modo aplicável do Outlook](https://docs.microsoft.com/outlook/add-ins/#extension-points)| Leitura|

##### <a name="example"></a>Exemplo

```
var internetMessageId = Office.context.mailbox.item.internetMessageId;
```

#### <a name="itemclass-string"></a>itemClass :String

Obtém a classe do item dos Serviços Web do Exchange do item selecionado. Somente modo de leitura.

A propriedade `itemClass` especifica a classe de mensagens do item selecionado. A seguir, estão as classes de mensagens padrão para o item de mensagem ou de compromisso.

| Tipo | Descrição | item class |
| --- | --- | --- |
| Itens de compromisso | São itens de calendário da classe de item `IPM.Appointment` ou `IPM.Appointment.Occurence`. | `IPM.Appointment`<br />`IPM.Appointment.Occurence` |
| Itens de mensagem | Incluem mensagens de e-mail que têm a classe de mensagem padrão `IPM.Note` e solicitações de reunião, respostas e cancelamentos que utilizam `IPM.Schedule.Meeting` como a classe de mensagem básica. | `IPM.Note`<br />`IPM.Schedule.Meeting.Request`<br />`IPM.Schedule.Meeting.Neg`<br />`IPM.Schedule.Meeting.Pos`<br />`IPM.Schedule.Meeting.Tent`<br />`IPM.Schedule.Meeting.Canceled` |

Você pode criar classes de mensagem personalizadas que ampliam uma classe de mensagens padrão, como, por exemplo, uma classe de mensagens de compromisso personalizada `IPM.Appointment.Contoso` .

##### <a name="type"></a>Tipo:

*   String

##### <a name="requirements"></a>Requisitos

|Requisito| Valor|
|---|---|
|[Versão mínima do conjunto de requisitos de caixa de correio](/javascript/office/requirement-sets/outlook-api-requirement-sets)| 1.0|
|[Nível de permissão mínimo](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| ReadItem|
|[Modo aplicável do Outlook](https://docs.microsoft.com/outlook/add-ins/#extension-points)| Leitura|

##### <a name="example"></a>Exemplo

```
var itemClass = Office.context.mailbox.item.itemClass;
```

#### <a name="nullable-itemid-string"></a>(nullable) itemId :String

Obtém o identificador do item dos Serviços Web do Exchange para o item atual. Somente modo de leitura.

> [!NOTE]
> O identificador retornado pela propriedade `itemId` é o mesmo que o identificador do item dos Serviços Web do Exchange. A propriedade `itemId` não é idêntica à ID de entrada do Outlook ou à ID usada pela API REST do Outlook. Antes de fazer chamadas de API REST usando esse valor, ele deve ser convertido usando [Office.context.mailbox.convertToRestId](office.context.mailbox.md#converttorestiditemid-restversion--string). Para mais informações, confira [Use as APIs REST do Outlook de um suplemento do Outlook](https://docs.microsoft.com/outlook/add-ins/use-rest-api#get-the-item-id).

A propriedade `itemId` não está disponível no modo de redação. Se o identificador de um item for obrigatório, o método [`saveAsync`](#saveasyncoptions-callback) pode ser usado para salvar o item no repositório, que retornará o identificador do item no parâmetro [`AsyncResult.value`](/javascript/api/office/office.asyncresult) na função de retorno de chamada.

##### <a name="type"></a>Tipo:

*   String

##### <a name="requirements"></a>Requisitos

|Requisito| Valor|
|---|---|
|[Versão mínima do conjunto de requisitos de caixa de correio](/javascript/office/requirement-sets/outlook-api-requirement-sets)| 1.0|
|[Nível de permissão mínimo](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| ReadItem|
|[Modo aplicável do Outlook](https://docs.microsoft.com/outlook/add-ins/#extension-points)| Leitura|

##### <a name="example"></a>Exemplo

O código a seguir verifica a presença de um identificador de item. Se a propriedade `itemId` retorna `null` ou `undefined`, ela salva o item no repositório e obtém o identificador do item do resultado assíncrono.

```
var itemId = Office.context.mailbox.item.itemId;
if (itemId === null || itemId == undefined) {
  Office.context.mailbox.item.saveAsync(function(result){
    itemId = result.value;
  });
}
```

####  <a name="itemtype-officemailboxenumsitemtypejavascriptapioutlook15officemailboxenumsitemtype"></a>itemType :[Office.MailboxEnums.ItemType](/javascript/api/outlook_1_5/office.mailboxenums.itemtype)

Obtém o tipo de item que uma instância representa.

A propriedade `itemType` retorna um dos valores de enumeração `ItemType`, indicando se a instância do objeto `item` é uma mensagem ou um compromisso.

##### <a name="type"></a>Tipo:

*   [Office.MailboxEnums.ItemType](/javascript/api/outlook_1_5/office.mailboxenums.itemtype)

##### <a name="requirements"></a>Requisitos

|Requisito| Valor|
|---|---|
|[Versão mínima do conjunto de requisitos de caixa de correio](/javascript/office/requirement-sets/outlook-api-requirement-sets)| 1.0|
|[Nível de permissão mínimo](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| ReadItem|
|[Modo aplicável do Outlook](https://docs.microsoft.com/outlook/add-ins/#extension-points)| Redigir ou ler|

##### <a name="example"></a>Exemplo

```
if (Office.context.mailbox.item.itemType == Office.MailboxEnums.ItemType.Message)
  // do something
else
  // do something else
```

####  <a name="location-stringlocationjavascriptapioutlook15officelocation"></a>location :String|[Location](/javascript/api/outlook_1_5/office.location)

Obtém ou define o local de um compromisso.

##### <a name="read-mode"></a>Modo de leitura

A propriedade `location` retorna uma sequência de caracteres que contém o local do compromisso.

##### <a name="compose-mode"></a>Modo de redação

A propriedade `location` retorna um objeto `Location` que fornece métodos para obter e definir o local do compromisso.

##### <a name="type"></a>Tipo:

*   String | [Location](/javascript/api/outlook_1_5/office.location)

##### <a name="requirements"></a>Requisitos

|Requisito| Valor|
|---|---|
|[Versão mínima do conjunto de requisitos de caixa de correio](/javascript/office/requirement-sets/outlook-api-requirement-sets)| 1.0|
|[Nível de permissão mínimo](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| ReadItem|
|[Modo aplicável do Outlook](https://docs.microsoft.com/outlook/add-ins/#extension-points)| Redigir ou ler|

##### <a name="example"></a>Exemplo

```
var userContext = { value : 1 };
Office.context.mailbox.item.location.getAsync( { context: userContext}, callback);

function callback(asyncResult) {
  var context = asyncResult.context;
  var location = asyncResult.value;
}
```

#### <a name="normalizedsubject-string"></a>normalizedSubject :String

Obtém o assunto de um item, com todos os prefixos removidos (incluindo `RE:` e `FWD:`). Somente modo de leitura.

A propriedade normalizedSubject obtém o assunto do item, com quaisquer prefixos padrão (como `RE:` e `FW:`) adicionados por programas de e-mail. Para obter o assunto do item com os prefixos intactos, use a propriedade [`subject`](#subject-stringsubjectjavascriptapioutlook15officesubject).

##### <a name="type"></a>Tipo:

*   Sequência de caracteres

##### <a name="requirements"></a>Requisitos

|Requisito| Valor|
|---|---|
|[Versão mínima do conjunto de requisitos de caixa de correio](/javascript/office/requirement-sets/outlook-api-requirement-sets)| 1.0|
|[Nível de permissão mínimo](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| ReadItem|
|[Modo aplicável do Outlook](https://docs.microsoft.com/outlook/add-ins/#extension-points)| Leitura|

##### <a name="example"></a>Exemplo

```
var normalizedSubject = Office.context.mailbox.item.normalizedSubject;
```

####  <a name="notificationmessages-notificationmessagesjavascriptapioutlook15officenotificationmessages"></a>notificationMessages :[NotificationMessages](/javascript/api/outlook_1_5/office.notificationmessages)

Obtém as mensagens de notificação de um item.

##### <a name="type"></a>Tipo:

*   [NotificationMessages](/javascript/api/outlook_1_5/office.notificationmessages)

##### <a name="requirements"></a>Requisitos

|Requisito| Valor|
|---|---|
|[Versão mínima do conjunto de requisitos de caixa de correio](/javascript/office/requirement-sets/outlook-api-requirement-sets)| 1.3|
|[Nível de permissão mínimo](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| ReadItem|
|[Modo aplicável do Outlook](https://docs.microsoft.com/outlook/add-ins/#extension-points)| Redigir ou ler|

####  <a name="optionalattendees-arrayemailaddressdetailsjavascriptapioutlook15officeemailaddressdetailsrecipientsjavascriptapioutlook15officerecipients"></a>optionalAttendees :Array.<[EmailAddressDetails](/javascript/api/outlook_1_5/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook_1_5/office.recipients)

Fornece acesso aos participantes opcionais de um evento. O tipo de objeto e o nível de acesso dependem do modo do item atual.

##### <a name="read-mode"></a>Modo de leitura

A propriedade `optionalAttendees` retorna uma matriz que contém um objeto `EmailAddressDetails` para cada participante opcional da reunião.

##### <a name="compose-mode"></a>Modo de redação

A propriedade `optionalAttendees` retorna um objeto `Recipients` que fornece métodos para obter ou atualizar os participantes opcionais de uma reunião.

##### <a name="type"></a>Tipo:

*   Array.<[EmailAddressDetails](/javascript/api/outlook_1_5/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook_1_5/office.recipients)

##### <a name="requirements"></a>Requisitos

|Requisito| Valor|
|---|---|
|[Versão mínima do conjunto de requisitos de caixa de correio](/javascript/office/requirement-sets/outlook-api-requirement-sets)| 1.0|
|[Nível de permissão mínimo](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| ReadItem|
|[Modo aplicável do Outlook](https://docs.microsoft.com/outlook/add-ins/#extension-points)| Redigir ou ler|

##### <a name="example"></a>Exemplo

```
Office.context.mailbox.item.optionalAttendees.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.optionalAttendees.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.optionalAttendees.getAsync(callback);

function callback(asyncResult) {
  var arrayOfOptionalAttendeesRecipients = asyncResult.value;
}
```

#### <a name="organizer-emailaddressdetailsjavascriptapioutlook15officeemailaddressdetails"></a>organizer :[EmailAddressDetails](/javascript/api/outlook_1_5/office.emailaddressdetails)

Obtém o endereço de e-mail do organizador da reunião para uma reunião especificada. Somente modo de leitura.

##### <a name="type"></a>Tipo:

*   [EmailAddressDetails](/javascript/api/outlook_1_5/office.emailaddressdetails)

##### <a name="requirements"></a>Requisitos

|Requisito| Valor|
|---|---|
|[Versão mínima do conjunto de requisitos de caixa de correio](/javascript/office/requirement-sets/outlook-api-requirement-sets)| 1.0|
|[Nível de permissão mínimo](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| ReadItem|
|[Modo aplicável do Outlook](https://docs.microsoft.com/outlook/add-ins/#extension-points)| Leitura|

##### <a name="example"></a>Exemplo

```
var organizerName = Office.context.mailbox.item.organizer.displayName;
var organizerAddress = Office.context.mailbox.item.organizer.emailAddress;
```

####  <a name="requiredattendees-arrayemailaddressdetailsjavascriptapioutlook15officeemailaddressdetailsrecipientsjavascriptapioutlook15officerecipients"></a>requiredAttendees :Array.<[EmailAddressDetails](/javascript/api/outlook_1_5/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook_1_5/office.recipients)

Fornece acesso aos participantes obrigatórios de um evento. O tipo de objeto e o nível de acesso dependem do modo do item atual.

##### <a name="read-mode"></a>Modo de leitura

A propriedade `requiredAttendees` retorna uma matriz que contém um objeto `EmailAddressDetails` para cada participante obrigatório da reunião.

##### <a name="compose-mode"></a>Modo de redação

A propriedade `requiredAttendees` retorna um objeto `Recipients` que fornece métodos para obter ou atualizar os participantes obrigatórios de uma reunião.

##### <a name="type"></a>Tipo:

*   Array.<[EmailAddressDetails](/javascript/api/outlook_1_5/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook_1_5/office.recipients)

##### <a name="requirements"></a>Requisitos

|Requisito| Valor|
|---|---|
|[Versão mínima do conjunto de requisitos de caixa de correio](/javascript/office/requirement-sets/outlook-api-requirement-sets)| 1.0|
|[Nível de permissão mínimo](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| ReadItem|
|[Modo aplicável do Outlook](https://docs.microsoft.com/outlook/add-ins/#extension-points)| Redigir ou ler|

##### <a name="example"></a>Exemplo

```
Office.context.mailbox.item.requiredAttendees.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.requiredAttendees.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.requiredAttendees.getAsync(callback);

function callback(asyncResult) {
  var arrayOfRequiredAttendeesRecipients = asyncResult.value;
}
```

#### <a name="sender-emailaddressdetailsjavascriptapioutlook15officeemailaddressdetails"></a>sender :[EmailAddressDetails](/javascript/api/outlook_1_5/office.emailaddressdetails)

Obtém o endereço de e-mail do remetente de uma mensagem de e-mail. Somente modo de leitura.

As propriedades [`from`](#from-emailaddressdetailsjavascriptapioutlook15officeemailaddressdetails) e `sender` correspondem a mesma pessoa, a menos que a mensagem seja enviada por um representante. Nesse caso, a propriedade `from` corresponde ao representado e a propriedade sender ao representante.

> [!NOTE]
> A propriedade `recipientType` do objeto `EmailAddressDetails` na propriedade `sender` é `undefined`.

##### <a name="type"></a>Tipo:

*   [EmailAddressDetails](/javascript/api/outlook_1_5/office.emailaddressdetails)

##### <a name="requirements"></a>Requisitos

|Requisito| Valor|
|---|---|
|[Versão mínima do conjunto de requisitos de caixa de correio](/javascript/office/requirement-sets/outlook-api-requirement-sets)| 1.0|
|[Nível de permissão mínimo](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| ReadItem|
|[Modo aplicável do Outlook](https://docs.microsoft.com/outlook/add-ins/#extension-points)| Leitura|

##### <a name="example"></a>Exemplo

```
var senderName = Office.context.mailbox.item.sender.displayName;
var senderAddress = Office.context.mailbox.item.sender.emailAddress;
```

####  <a name="start-datetimejavascriptapioutlook15officetime"></a>start :Date|[Time](/javascript/api/outlook_1_5/office.time)

Obtém ou define a data e a hora em que o compromisso deve começar.

A propriedade `start` é expressa como um valor de data e hora no Tempo Universal Coordenado (UTC). Você pode usar o método [`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttimejavascriptapioutlook15officelocalclienttime) para converter o valor para a data e a hora local do cliente.

##### <a name="read-mode"></a>Modo de leitura

A propriedade `start` retorna um objeto `Date`.

##### <a name="compose-mode"></a>Modo de redação

A propriedade `start` retorna um objeto `Time`.

Ao usar o método [`Time.setAsync`](/javascript/api/outlook_1_5/office.time#setasync-datetime--options--callback-) para definir a hora de início, deve-se usar o método [`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) para converter a hora local no cliente para UTC para o servidor.

##### <a name="type"></a>Tipo:

*   Date | [Time](/javascript/api/outlook_1_5/office.time)

##### <a name="requirements"></a>Requisitos

|Requisito| Valor|
|---|---|
|[Versão mínima do conjunto de requisitos de caixa de correio](/javascript/office/requirement-sets/outlook-api-requirement-sets)| 1.0|
|[Nível de permissão mínimo](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| ReadItem|
|[Modo aplicável do Outlook](https://docs.microsoft.com/outlook/add-ins/#extension-points)| Redigir ou ler|

##### <a name="example"></a>Exemplo

O exemplo a seguir define a hora de início de um compromisso no modo de redação usando o método [`setAsync`](/javascript/api/outlook_1_5/office.time#setasync-datetime--options--callback-) do objeto `Time`.

```
var startTime = new Date("3/14/2015");
var options = {
  // Pass information that can be used
  // in the callback
     asyncContext: {verb:"Set"}
}
Office.context.mailbox.item.start.setAsync(startTime, options, function(result) {
  if (result.error) {
    console.debug(result.error);
  } else {
    // Access the asyncContext that was passed to the setAsync function
    console.debug("Start Time " + result.asyncContext.verb);
  }
});
```

####  <a name="subject-stringsubjectjavascriptapioutlook15officesubject"></a>subject :String|[Subject](/javascript/api/outlook_1_5/office.subject)

Obtém ou define a descrição que aparece no campo de assunto de um item.

A propriedade `subject` obtém ou define o assunto completo do item, conforme enviado pelo servidor de e-mail.

##### <a name="read-mode"></a>Modo de leitura

A propriedade `subject` retorna uma sequência de caracteres. Use a propriedade [`normalizedSubject`](#normalizedsubject-string) para obter o assunto sem os prefixos iniciais, como `RE:` e `FW:`.

```
var subject = Office.context.mailbox.item.subject;
```

##### <a name="compose-mode"></a>Modo de redação

A propriedade `subject` retorna um objeto `Subject` que fornece métodos para obter e definir o assunto.

```
Office.context.mailbox.item.subject.getAsync(callback);

function callback(asyncResult) {
  var subject = asyncResult.value;
}
```

##### <a name="type"></a>Tipo:

*   String | [Subject](/javascript/api/outlook_1_5/office.subject)

##### <a name="requirements"></a>Requisitos

|Requisito| Valor|
|---|---|
|[Versão mínima do conjunto de requisitos de caixa de correio](/javascript/office/requirement-sets/outlook-api-requirement-sets)| 1.0|
|[Nível de permissão mínimo](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| ReadItem|
|[Modo aplicável do Outlook](https://docs.microsoft.com/outlook/add-ins/#extension-points)| Redigir ou ler|

####  <a name="to-arrayemailaddressdetailsjavascriptapioutlook15officeemailaddressdetailsrecipientsjavascriptapioutlook15officerecipients"></a>to :Array.<[EmailAddressDetails](/javascript/api/outlook_1_5/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook_1_5/office.recipients)

Fornece acesso aos destinatários na linha **Para** de uma mensagem. O tipo de objeto e o nível de acesso dependem do modo do item atual.

##### <a name="read-mode"></a>Modo de leitura

A propriedade `to` retorna uma matriz que contém um objeto `EmailAddressDetails` para cada destinatário listado na linha **Para** da mensagem. A coleção é limitada a um número máximo de 100 membros.

##### <a name="compose-mode"></a>Modo de redação

A propriedade `to` retorna um objeto `Recipients` que fornece métodos para obter ou atualizar os destinatários na linha  **Para** da mensagem.

##### <a name="type"></a>Tipo:

*   Array.<[EmailAddressDetails](/javascript/api/outlook_1_5/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook_1_5/office.recipients)

##### <a name="requirements"></a>Requisitos

|Requisito| Valor|
|---|---|
|[Versão mínima do conjunto de requisitos de caixa de correio](/javascript/office/requirement-sets/outlook-api-requirement-sets)| 1.0|
|[Nível de permissão mínimo](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| ReadItem|
|[Modo aplicável do Outlook](https://docs.microsoft.com/outlook/add-ins/#extension-points)| Redigir ou ler|

##### <a name="example"></a>Exemplo

```
Office.context.mailbox.item.to.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.to.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.to.getAsync(callback);

function callback(asyncResult) {
  var arrayOfToRecipients = asyncResult.value;
}
```

### <a name="methods"></a>Métodos

####  <a name="addfileattachmentasyncuri-attachmentname-options-callback"></a>addFileAttachmentAsync(uri, attachmentName, [options], [callback])

Adiciona um arquivo a uma mensagem ou a um compromisso em forma de anexo.

O método `addFileAttachmentAsync` carrega o arquivo no URI especificado e anexa-o ao item no formulário de redação.

Posteriormente, você poderá usar o identificador com o método [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) para remover o anexo na mesma sessão.

##### <a name="parameters"></a>Parâmetros:

|Nome| Tipo| Atributos| Descrição|
|---|---|---|---|
|`uri`| String||O URI que fornece a localização do arquivo anexado à mensagem ou ao compromisso. O comprimento máximo é de 2048 caracteres.|
|`attachmentName`| String||O nome do anexo que é mostrado enquanto o anexo está sendo carregado. O comprimento máximo é de 255 caracteres.|
|`options`| Object| &lt;opcional&gt;|Um literal de objeto que contém uma ou mais das propriedades a seguir.|
| `options.asyncContext` | Object | &lt;opcional&gt; | Os desenvolvedores podem fornecer qualquer objeto que desejarem acessar no método de retorno de chamada. |
| `options.isInline` | Booleano | &lt;opcional&gt; | Se for `true`, indicará que o anexo será mostrado em linha no corpo da mensagem e não deverá ser exibido na lista de anexos. |
|`callback`| function| &lt;opcional&gt;|Quando o método for concluído, a função passada no parâmetro `callback` é chamada com um parâmetro único, `asyncResult`, que é um objeto [`AsyncResult`](/javascript/api/office/office.asyncresult). <br/>Em caso de êxito, o identificador do anexo será fornecido na propriedade `asyncResult.value`.<br/>Se houver falha ao carregar o anexo, o objeto `asyncResult` conterá um objeto `Error` que fornecerá uma descrição do erro.|

##### <a name="errors"></a>Erros

| Código de erro | Descrição |
|------------|-------------|
| `AttachmentSizeExceeded` | O anexo é maior do que permitido. |
| `FileTypeNotSupported` | O anexo tem uma extensão que não é permitida. |
| `NumberOfAttachmentsExceeded` | A mensagem ou o compromisso tem muitos anexos. |

##### <a name="requirements"></a>Requisitos

|Requisito| Valor|
|---|---|
|[Versão mínima do conjunto de requisitos de caixa de correio](/javascript/office/requirement-sets/outlook-api-requirement-sets)| 1.1|
|[Nível de permissão mínimo](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| ReadWriteItem|
|[Modo aplicável do Outlook](https://docs.microsoft.com/outlook/add-ins/#extension-points)| Redigir|

##### <a name="examples"></a>Exemplos

```js
function callback(result) {
  if (result.error) {
    showMessage(result.error);
  } else {
    showMessage("Attachment added");
  }
}

function addAttachment() {
  // The values in asyncContext can be accessed in the callback
  var options = { 'asyncContext': { var1: 1, var2: 2 } };

  var attachmentURL = "https://contoso.com/rtm/icon.png";
  Office.context.mailbox.item.addFileAttachmentAsync(attachmentURL, attachmentURL, options, callback);
}
```

O exemplo a seguir adiciona um arquivo de imagem como um anexo em linha e faz referência ao anexo no corpo da mensagem.

```js
Office.context.mailbox.item.addFileAttachmentAsync
(
  "http://i.imgur.com/WJXklif.png",
  "cute_bird.png",
  {
    isInline: true
  },
  function (asyncResult) {
    Office.context.mailbox.item.body.setAsync(
      "<p>Here's a cute bird!</p><img src='cid:cute_bird.png'>",
      {
        "coercionType": "html"
      },
      function (asyncResult) {
        
      }
    );
  }
);
```

####  <a name="additemattachmentasyncitemid-attachmentname-options-callback"></a>addItemAttachmentAsync(itemId, attachmentName, [options], [callback])

Adiciona um item do Exchange, como uma mensagem, em forma de anexo na mensagem ou no compromisso.

O método `addItemAttachmentAsync` anexa o item com o identificador especificado  do Exchange ao item no formulário de redação. Se você especificar um método de retorno de chamada, o método será chamado com um parâmetro  `asyncResult` que contém o identificador do anexo ou um código que indica qualquer erro que tenha ocorrido ao anexar o item. Você pode usar o parâmetro `options` para passar informações de estado ao método de retorno de chamada, se necessário.

Posteriormente, você poderá usar o identificador com o método [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) para remover o anexo na mesma sessão.

Se o suplemento do Office estiver em execução no Outlook Web App, o método `addItemAttachmentAsync` pode anexar itens aos itens que não sejam aqueles que você esteja editando. No entanto, isso não é suportado e não é recomendado.

##### <a name="parameters"></a>Parâmetros:

|Nome| Tipo| Attributes| Descrição|
|---|---|---|---|
|`itemId`| String||O identificador do Exchange do item a ser anexado. O comprimento máximo é de 100 caracteres.|
|`attachmentName`| String||O assunto do item a ser anexado. O comprimento máximo é de 255 caracteres.|
|`options`| Object| &lt;opcional&gt;|Um literal de objeto que contém uma ou mais das propriedades a seguir.|
|`options.asyncContext`| Object| &lt;opcional&gt;|Os desenvolvedores podem fornecer qualquer objeto que desejarem acessar no método de retorno de chamada.|
|`callback`| function| &lt;opcional&gt;|Quando o método for concluído, a função passada no parâmetro `callback` é chamada com um parâmetro único, `asyncResult`, que é um objeto [`AsyncResult`](/javascript/api/office/office.asyncresult). <br/>Em caso de êxito, o identificador do anexo será fornecido na propriedade `asyncResult.value`.<br/>Se houver falha ao adicionar o anexo, o objeto `asyncResult` conterá um objeto `Error` que fornecerá uma descrição do erro.|

##### <a name="errors"></a>Erros

| Código de erro | Descrição |
|------------|-------------|
| `NumberOfAttachmentsExceeded` | A mensagem ou o compromisso tem muitos anexos. |

##### <a name="requirements"></a>Requisitos

|Requisito| Valor|
|---|---|
|[Versão mínima do conjunto de requisitos de caixa de correio](/javascript/office/requirement-sets/outlook-api-requirement-sets)| 1.1|
|[Nível de permissão mínimo](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| ReadWriteItem|
|[Modo aplicável do Outlook](https://docs.microsoft.com/outlook/add-ins/#extension-points)| Redigir|

##### <a name="example"></a>Exemplo

O exemplo a seguir adiciona um item existente do Outlook como um anexo com o nome `My Attachment`.

```
function callback(result) {
  if (result.error) {
    showMessage(result.error);
  } else {
    showMessage("Attachment added");
  }
}

function addAttachment() {
  // EWS ID of item to attach
  // (Shortened for readability)
  var itemId = "AAMkADI1...AAA=";

  // The values in asyncContext can be accessed in the callback
  var options = { 'asyncContext': { var1: 1, var2: 2 } };

  Office.context.mailbox.item.addItemAttachmentAsync(itemId, "My Attachment", options, callback);
}
```

####  <a name="close"></a>close()

Fecha o item atual que está sendo redigido.

O comportamento do método `close` depende do estado atual do item que está sendo redigido. Se o item possuir alterações não salvas, o cliente solicitará que o usuário salve, descarte ou cancele a ação ao fechar.

> [!NOTE]
> No Outlook na Web, se o item for um compromisso e tiver sido salvo anteriormente usando `saveAsync`, será solicitado ao usuário para salvar, descartar ou cancelar, mesmo se nenhuma alteração tenha ocorrido após o item ter sido salvo pela última vez.

No cliente do Outlook para área de trabalho, se a mensagem for uma resposta embutida, o método `close` não terá efeito.

##### <a name="requirements"></a>Requisitos

|Requisito| Valor|
|---|---|
|[Versão mínima do conjunto de requisitos de caixa de correio](/javascript/office/requirement-sets/outlook-api-requirement-sets)| 1.3|
|[Nível de permissão mínimo](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| Restrito|
|[Modo aplicável do Outlook](https://docs.microsoft.com/outlook/add-ins/#extension-points)| Redigir|

#### <a name="displayreplyallformformdata"></a>displayReplyAllForm(formData)

Exibe um formulário de resposta que inclui o remetente e todos os destinatários da mensagem selecionada ou o organizador e todos os participantes do compromisso selecionado.

> [!NOTE]
> Esse método não é compatível com o Outlook para iOS ou o Outlook para Android.

No Outlook Web App, o formulário de resposta é exibido como um formulário pop-out no modo de exibição de três colunas e um formulário pop-up no modo de exibição de uma ou duas colunas.

Se qualquer um dos parâmetros da sequência de caracteres exceder o limite, `displayReplyAllForm` gerará uma exceção.

Quando os anexos são especificados no parâmetro `formData.attachments`, o Outlook e o Outlook Web App tentam baixar todos os anexos e anexá-los ao formulário de resposta. Se a adição de anexos falhar, será exibido um erro na interface de usuário do formulário. Se isso não for possível, nenhuma mensagem de erro será apresentada.

##### <a name="parameters"></a>Parâmetros:

| Nome | Tipo | Attributes | Descrição |
|---|---|---|---|
|`formData`| String | Object| |Uma sequência de caracteres que contém texto e HTML e que representa o corpo do formulário de resposta. A sequência de caracteres está limitada a 32 KB.<br/>**OR**<br/>Um objeto que contém os dados do corpo ou do anexo e uma função de retorno de chamada. O objeto é definido da seguinte maneira. |
| `formData.htmlBody` | String | &lt;opcional&gt; | Uma sequência de caracteres que contém texto e HTML e que representa o corpo do formulário de resposta. A sequência de caracteres está limitada a 32 KB.
| `formData.attachments` | Array.&lt;Object&gt; | &lt;opcional&gt; | Uma matriz de objetos JSON que são anexos do arquivo ou do item. |
| `formData.attachments.type` | String | | Indica o tipo de anexo. Deve ser `file` para um anexo de arquivo ou `item` para um anexo de item. |
| `formData.attachments.name` | String | | Uma sequência de caracteres que contém o nome do anexo, com até 255 caracteres de comprimento.|
| `formData.attachments.url` | String | | Usado somente se `type` estiver definido como `file`. O URI do local para o arquivo. |
| `formData.attachments.isInline` | Booleano | | Usado somente se `type` estiver definido como `file`. Se for `true`, indicará que o anexo será mostrado em linha no corpo da mensagem e não deverá ser exibido na lista de anexos. |
| `formData.attachments.itemId` | String | | Usado somente se `type` estiver definido como `item`. A ID do item do EWS do anexo. Isso é uma sequência de caracteres de até 100 caracteres. |
| `callback` | function | &lt;opcional&gt; | Quando o método for concluído, a função passada ao parâmetro `callback` será chamada com um parâmetro único `asyncResult`, que é um objeto [AsyncResult](/javascript/api/office/office.asyncresult). |

##### <a name="requirements"></a>Requisitos

|Requisito| Valor|
|---|---|
|[Versão mínima do conjunto de requisitos de caixa de correio](/javascript/office/requirement-sets/outlook-api-requirement-sets)| 1.0|
|[Nível de permissão mínimo](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| ReadItem|
|[Modo aplicável do Outlook](https://docs.microsoft.com/outlook/add-ins/#extension-points)| Leitura|

##### <a name="examples"></a>Exemplos

O código a seguir transmite uma sequência de caracteres para a função `displayReplyAllForm`.

```
Office.context.mailbox.item.displayReplyAllForm('hello there');
Office.context.mailbox.item.displayReplyAllForm('<b>hello there</b>');
```

Responder com um corpo vazio.

```
Office.context.mailbox.item.displayReplyAllForm({});
```

Responder apenas com um corpo.

```
Office.context.mailbox.item.displayReplyAllForm(
{
  'htmlBody' : 'hi'
});
```

Responder com um corpo e um anexo de arquivo.

```
Office.context.mailbox.item.displayReplyAllForm(
{
  'htmlBody' : 'hi',
  'attachments' :
  [
    {
      'type' : Office.MailboxEnums.AttachmentType.File,
      'name' : 'squirrel.png',
      'url' : 'http://i.imgur.com/sRgTlGR.jpg'
    }
  ]
});
```

Responder com um corpo e um anexo de item.

```
Office.context.mailbox.item.displayReplyAllForm(
{
  'htmlBody' : 'hi',
  'attachments' :
  [
    {
      'type' : 'item',
      'name' : 'rand',
      'itemId' : Office.context.mailbox.item.itemId
    }
  ]
});
```

Responder com um corpo, um anexo de arquivo, um anexo de item e um retorno de chamada.

```
Office.context.mailbox.item.displayReplyAllForm(
{
  'htmlBody' : 'hi',
  'attachments' :
  [
    {
      'type' : Office.MailboxEnums.AttachmentType.File,
      'name' : 'squirrel.png',
      'url' : 'http://i.imgur.com/sRgTlGR.jpg'
    },
    {
      'type' : 'item',
      'name' : 'rand',
      'itemId' : Office.context.mailbox.item.itemId
    }
  ],
  'callback' : function(asyncResult)
  {
    console.log(asyncResult.value);
  }
});
```

#### <a name="displayreplyformformdata"></a>displayReplyForm(formData)

Exibe um formulário de resposta que inclui o remetente da mensagem selecionada ou o organizador do compromisso selecionado.

> [!NOTE]
> Esse método não é compatível com o Outlook para iOS ou o Outlook para Android.

No Outlook Web App, o formulário de resposta é exibido como um formulário pop-out no modo de exibição de três colunas e um formulário pop-up no modo de exibição de uma ou duas colunas.

Se qualquer um dos parâmetros da sequência de caracteres exceder o limite, `displayReplyForm` gerará uma exceção.

Quando os anexos são especificados no parâmetro `formData.attachments`, o Outlook e o Outlook Web App tentam baixar todos os anexos e anexá-los ao formulário de resposta. Se a adição de anexos falhar, será exibido um erro na interface de usuário do formulário. Se isso não for possível, nenhuma mensagem de erro será apresentada.

##### <a name="parameters"></a>Parâmetros:

| Nome | Tipo | Attributes | Descrição |
|---|---|---|---|
|`formData`| String | Object| | Uma sequência de caracteres que contém texto e HTML e que representa o corpo do formulário de resposta. A sequência de caracteres está limitada a 32 KB.<br/>**OR**<br/>Um objeto que contém os dados do corpo ou do anexo e uma função de retorno de chamada. O objeto é definido da seguinte maneira. |
| `formData.htmlBody` | String | &lt;opcional&gt; | Uma sequência de caracteres que contém texto e HTML e que representa o corpo do formulário de resposta. A sequência de caracteres está limitada a 32 KB.
| `formData.attachments` | Array.&lt;Object&gt; | &lt;opcional&gt; | Uma matriz de objetos JSON que são anexos do arquivo ou do item. |
| `formData.attachments.type` | String | | Indica o tipo de anexo. Deve ser `file` para um anexo de arquivo ou `item` para um anexo de item. |
| `formData.attachments.name` | String | | Uma sequência de caracteres que contém o nome do anexo, com até 255 caracteres de comprimento.|
| `formData.attachments.url` | String | | Usado somente se `type` estiver definido como `file`. O URI do local para o arquivo. |
| `formData.attachments.isInline` | Booleano | | Usado somente se `type` estiver definido como `file`. Se for `true`, indicará que o anexo será mostrado em linha no corpo da mensagem e não deverá ser exibido na lista de anexos. |
| `formData.attachments.itemId` | String | | Usado somente se `type` estiver definido como `item`. A ID do item do EWS do anexo. Isso é uma sequência de caracteres de até 100 caracteres. |
| `callback` | function | &lt;opcional&gt; | Quando o método for concluído, a função passada ao parâmetro `callback` será chamada com um parâmetro único `asyncResult`, que é um objeto [AsyncResult](/javascript/api/office/office.asyncresult). |

##### <a name="requirements"></a>Requisitos

|Requisito| Valor|
|---|---|
|[Versão mínima do conjunto de requisitos de caixa de correio](/javascript/office/requirement-sets/outlook-api-requirement-sets)| 1.0|
|[Nível de permissão mínimo](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| ReadItem|
|[Modo aplicável do Outlook](https://docs.microsoft.com/outlook/add-ins/#extension-points)| Leitura|

##### <a name="examples"></a>Exemplos

O código a seguir transmite uma sequência de caracteres para a função `displayReplyForm`.

```
Office.context.mailbox.item.displayReplyForm('hello there');
Office.context.mailbox.item.displayReplyForm('<b>hello there</b>');
```

Responder com um corpo vazio.

```
Office.context.mailbox.item.displayReplyForm({});
```

Responder apenas com um corpo.

```
Office.context.mailbox.item.displayReplyForm(
{
  'htmlBody' : 'hi'
});
```

Responder com um corpo e um anexo de arquivo.

```
Office.context.mailbox.item.displayReplyForm(
{
  'htmlBody' : 'hi',
  'attachments' :
  [
    {
      'type' : Office.MailboxEnums.AttachmentType.File,
      'name' : 'squirrel.png',
      'url' : 'http://i.imgur.com/sRgTlGR.jpg'
    }
  ]
});
```

Responder com um corpo e um anexo de item.

```
Office.context.mailbox.item.displayReplyForm(
{
  'htmlBody' : 'hi',
  'attachments' :
  [
    {
      'type' : 'item',
      'name' : 'rand',
      'itemId' : Office.context.mailbox.item.itemId
    }
  ]
});
```

Responder com um corpo, um anexo de arquivo, um anexo de item e um retorno de chamada.

```
Office.context.mailbox.item.displayReplyForm(
{
  'htmlBody' : 'hi',
  'attachments' :
  [
    {
      'type' : Office.MailboxEnums.AttachmentType.File,
      'name' : 'squirrel.png',
      'url' : 'http://i.imgur.com/sRgTlGR.jpg'
    },
    {
      'type' : 'item',
      'name' : 'rand',
      'itemId' : Office.context.mailbox.item.itemId
    }
  ],
  'callback' : function(asyncResult)
  {
    console.log(asyncResult.value);
  }
});
```

#### <a name="getentities--entitiesjavascriptapioutlook15officeentities"></a>getEntities() → {[Entities](/javascript/api/outlook_1_5/office.entities)}

Obtém as entidades encontradas no corpo do item selecionado.

> [!NOTE]
> Esse método não é compatível com o Outlook para iOS ou o Outlook para Android.

##### <a name="requirements"></a>Requisitos

|Requisito| Valor|
|---|---|
|[Versão mínima do conjunto de requisitos de caixa de correio](/javascript/office/requirement-sets/outlook-api-requirement-sets)| 1.0|
|[Nível de permissão mínimo](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| ReadItem|
|[Modo aplicável do Outlook](https://docs.microsoft.com/outlook/add-ins/#extension-points)| Leitura|

##### <a name="returns"></a>Retorna:

Tipo: [Entities](/javascript/api/outlook_1_5/office.entities)

##### <a name="example"></a>Exemplo

O exemplo a seguir acessa as entidades de contatos no corpo do item atual.

```
var contacts = Office.context.mailbox.item.getEntities().contacts;
```

#### <a name="getentitiesbytypeentitytype--nullable-arraystringcontactjavascriptapioutlook15officecontactmeetingsuggestionjavascriptapioutlook15officemeetingsuggestionphonenumberjavascriptapioutlook15officephonenumbertasksuggestionjavascriptapioutlook15officetasksuggestion"></a>getEntitiesByType(entityType) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook_1_5/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_5/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_5/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_5/office.tasksuggestion))>}

Obtém uma matriz de todas as entidades do tipo de entidade especificado encontrado no corpo do item selecionado.

> [!NOTE]
> Esse método não é compatível com o Outlook para iOS ou o Outlook para Android.

##### <a name="parameters"></a>Parâmetros:

|Nome| Tipo| Descrição|
|---|---|---|
|`entityType`| [Office.MailboxEnums.EntityType](/javascript/api/outlook_1_5/office.mailboxenums.entitytype)|Um dos valores de enumeração de EntityType.|

##### <a name="requirements"></a>Requisitos

|Requisito| Valor|
|---|---|
|[Versão mínima do conjunto de requisitos de caixa de correio](/javascript/office/requirement-sets/outlook-api-requirement-sets)| 1.0|
|[Nível de permissão mínimo](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| Restrito|
|[Modo aplicável do Outlook](https://docs.microsoft.com/outlook/add-ins/#extension-points)| Leitura|

##### <a name="returns"></a>Retorna:

Se o valor passado em `entityType` não for um membro válido da enumeração `EntityType`, o método retornará nulo. Se nenhuma entidade do tipo especificado estiver presente no corpo do item, o método retornará uma matriz vazia. Caso contrário, o tipo dos objetos na matriz retornada dependerá do tipo de entidade solicitado no parâmetro `entityType`.

Embora o nível de permissão mínimo para usar esse método seja **Restricted**, alguns tipos de entidade exigem a permissão **ReadItem** para obter acesso, conforme especificado na tabela a seguir.

| Valor de `entityType` | Tipo de objetos na matriz retornada | Nível de Permissão Exigido |
| --- | --- | --- |
| `Address` | String | **Restrito** |
| `Contact` | Contact | **ReadItem** |
| `EmailAddress` | String | **ReadItem** |
| `MeetingSuggestion` | MeetingSuggestion | **ReadItem** |
| `PhoneNumber` | PhoneNumber | **Restrito** |
| `TaskSuggestion` | TaskSuggestion | **ReadItem** |
| `URL` | String | **Restrito** |

Type: Array.<(String|[Contact](/javascript/api/outlook_1_5/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_5/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_5/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_5/office.tasksuggestion))>

##### <a name="example"></a>Exemplo

O exemplo a seguir mostra como acessar uma matriz de sequências de caracteres que representa endereços postais no corpo do item atual.

```
// The initialize function is required for all apps.
Office.initialize = function () {
  // Checks for the DOM to load using the jQuery ready function.
  $(document).ready(function () {
    // After the DOM is loaded, app-specific code can run.
    var item = Office.context.mailbox.item;
    // Get an array of strings that represent postal addresses in the current item's body.
    var addresses = item.getEntitiesByType(Office.MailboxEnums.EntityType.Address);
    // Continue processing the array of addresses.
  });
}
```

#### <a name="getfilteredentitiesbynamename--nullable-arraystringcontactjavascriptapioutlook15officecontactmeetingsuggestionjavascriptapioutlook15officemeetingsuggestionphonenumberjavascriptapioutlook15officephonenumbertasksuggestionjavascriptapioutlook15officetasksuggestion"></a>getFilteredEntitiesByName(name) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook_1_5/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_5/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_5/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_5/office.tasksuggestion))>}

Retorna entidades bem conhecidas no item selecionado que passam o filtro nomeado definido no arquivo de manifesto XML.

> [!NOTE]
> Esse método não é compatível com o Outlook para iOS ou o Outlook para Android.

O método `getFilteredEntitiesByName` retorna as entidades que correspondem à expressão regular definida no elemento de regra [ItemHasKnownEntity](/javascript/office/manifest/rule#itemhasknownentity-rule) no arquivo de manifesto XML com o valor do elemento `FilterName` especificado.

##### <a name="parameters"></a>Parâmetros:

|Nome| Tipo| Descrição|
|---|---|---|
|`name`| String|O nome do elemento de regra `ItemHasKnownEntity` que define o filtro a ser correspondido.|

##### <a name="requirements"></a>Requisitos

|Requisito| Valor|
|---|---|
|[Versão mínima do conjunto de requisitos de caixa de correio](/javascript/office/requirement-sets/outlook-api-requirement-sets)| 1.0|
|[Nível de permissão mínimo](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| ReadItem|
|[Modo aplicável do Outlook](https://docs.microsoft.com/outlook/add-ins/#extension-points)| Leitura|

##### <a name="returns"></a>Retorna:

Se não houver nenhum elemento `ItemHasKnownEntity` no manifesto com um valor de elemento `FilterName` que corresponda ao parâmetro `name`, o método retornará `null`. Se o parâmetro `name` corresponder a um elemento `ItemHasKnownEntity` no manifesto, mas não houver entidades no item atual que correspondam, o método retornará uma matriz vazia.

Type: Array.<(String|[Contact](/javascript/api/outlook_1_5/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_5/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_5/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_5/office.tasksuggestion))>

#### <a name="getregexmatches--object"></a>getRegExMatches() → {Object}

Retorna valores de sequência de caracteres ao item selecionado que correspondem às expressões regulares definidas no arquivo de manifesto XML.

> [!NOTE]
> Esse método não é compatível com o Outlook para iOS ou o Outlook para Android.

O método `getRegExMatches` retorna as sequências de caracteres que correspondem à expressão regular definida em cada elemento de regra `ItemHasRegularExpressionMatch` ou `ItemHasKnownEntity` no arquivo de manifesto XML. Para uma regra `ItemHasRegularExpressionMatch`, uma sequência de caracteres correspondente deve ocorrer na propriedade do item especificada por essa regra. O tipo simples `PropertyName` define as propriedades compatíveis.

Por exemplo, considere que um manifesto de suplemento tenha o seguinte elemento `Rule` :

```
<Rule xsi:type="RuleCollection" Mode="And">
  <Rule xsi:type="ItemIs" FormType="Read" ItemType="Message" />
  <Rule xsi:type="RuleCollection" Mode="Or">
    <Rule xsi:type="ItemHasRegularExpressionMatch" RegExName="fruits" RegExValue="apple|banana|coconut" PropertyName="BodyAsPlaintext" IgnoreCase="true" />
    <Rule xsi:type="ItemHasRegularExpressionMatch" RegExName="veggies" RegExValue="tomato|onion|spinach|broccoli" PropertyName="BodyAsPlaintext" IgnoreCase="true" />
  </Rule>
</Rule>
```

O objeto retornado por `getRegExMatches` teria duas propriedades: `fruits` e `veggies`.

```
{
  'fruits': ['apple','banana','Banana','coconut'],
  'veggies': ['tomato','onion','spinach','broccoli']
}
```

Se você especificar uma regra `ItemHasRegularExpressionMatch` na propriedade do corpo de um item, a expressão regular deverá filtrar ainda mais o corpo e não tentará retornar o corpo inteiro do item. Usar uma expressão regular como `.*` para obter o corpo inteiro de um item nem sempre retornará os resultados esperados. Em vez disso, use o método [`Body.getAsync`](/javascript/api/outlook_1_5/office.body#getasync-coerciontype--options--callback-) para recuperar o corpo inteiro.

##### <a name="requirements"></a>Requisitos

|Requisito| Valor|
|---|---|
|[Versão mínima do conjunto de requisitos de caixa de correio](/javascript/office/requirement-sets/outlook-api-requirement-sets)| 1.0|
|[Nível de permissão mínimo](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| ReadItem|
|[Modo aplicável do Outlook](https://docs.microsoft.com/outlook/add-ins/#extension-points)| Leitura|

##### <a name="returns"></a>Retorna:

Um objeto que contém matrizes de sequências de caracteres que correspondem às expressões regulares definidas no arquivo de manifesto XML. O nome de cada matriz é igual ao valor correspondente do atributo `RegExName` da regra `ItemHasRegularExpressionMatch` correspondente ou do atributo `FilterName` da regra `ItemHasKnownEntity` correspondente.

<dl class="param-type">

<dt>Tipo</dt>

<dd>Object</dd>

</dl>

##### <a name="example"></a>Exemplo

O exemplo a seguir mostra como acessar a matriz de correspondências para os <rule>elementos da expressão regular `fruits` e `veggies` que são especificados no manifesto.</rule>

```
var allMatches = Office.context.mailbox.item.getRegExMatches();
var fruits = allMatches.fruits;
var veges = allMatches.veggies;
```

#### <a name="getregexmatchesbynamename--nullable-array-string-"></a>getRegExMatchesByName(name) → (nullable) {Array.< String >}

Retorna valores de sequência de caracteres no item selecionado que correspondem à expressão regular nomeada definida no arquivo de manifesto XML.

> [!NOTE]
> Esse método não é compatível com o Outlook para iOS ou o Outlook para Android.

O método `getRegExMatchesByName` retorna as sequências de caracteres que correspondem à expressão regular definida no elemento de regra `ItemHasRegularExpressionMatch` no arquivo de manifesto XML com o valor de elemento `RegExName` especificado.

Se você especificar uma regra `ItemHasRegularExpressionMatch` na propriedade do corpo de um item, a expressão regular deverá filtrar mais o corpo e não tentar retornar todo o corpo do item. Usar uma expressão regular, como `.*`, para obter todo o corpo de um item nem sempre retorna os resultados esperados.

##### <a name="parameters"></a>Parâmetros:

|Nome| Tipo| Descrição|
|---|---|---|
|`name`| String|O nome do elemento de regra `ItemHasRegularExpressionMatch` que define o filtro a ser correspondido.|

##### <a name="requirements"></a>Requisitos

|Requisito| Valor|
|---|---|
|[Versão mínima do conjunto de requisitos de caixa de correio](/javascript/office/requirement-sets/outlook-api-requirement-sets)| 1.0|
|[Nível de permissão mínimo](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| ReadItem|
|[Modo aplicável do Outlook](https://docs.microsoft.com/outlook/add-ins/#extension-points)| Leitura|

##### <a name="returns"></a>Retorna:

Uma matriz que contém as sequências de caracteres que correspondem à expressão regular definida no arquivo de manifesto XML.

<dl class="param-type">

<dt>Tipo</dt>

<dd>Array.< String ></dd>

</dl>

##### <a name="example"></a>Exemplo

```
var fruits = Office.context.mailbox.item.getRegExMatchesByName("fruits");
var veggies = Office.context.mailbox.item.getRegExMatchesByName("veggies");
```

####  <a name="getselecteddataasynccoerciontype-options-callback--string"></a>getSelectedDataAsync(coercionType, [options], callback) → {String}

Retorna de forma assíncrona os dados selecionados do assunto ou do corpo de uma mensagem.

Se não houver seleção, mas o cursor estiver no corpo ou no assunto, o método retornará nulo para os dados selecionados. Se um campo que não seja o corpo ou o assunto estiver selecionado, o método retornará o erro `InvalidSelection`.

##### <a name="parameters"></a>Parâmetros:

|Nome| Tipo| Attributes| Descrição|
|---|---|---|---|
|`coercionType`| [Office.CoercionType](office.md#coerciontype-string)||Solicita um formato para os dados. Se for texto, o método retornará o texto sem formatação em forma de sequência de caracteres, removendo quaisquer tags HTML presentes. Se for HTML, o método retornará o texto selecionado, seja ele texto sem formatação ou HTML.|
|`options`| Object| &lt;opcional&gt;|Um literal de objeto que contém uma ou mais das propriedades a seguir.|
|`options.asyncContext`| Object| &lt;opcional&gt;|Os desenvolvedores podem fornecer qualquer objeto que desejarem acessar no método de retorno de chamada.|
|`callback`| function||Quando o método for concluído, a função passada no parâmetro `callback` é chamada com um parâmetro único, `asyncResult`, que é um objeto [`AsyncResult`](/javascript/api/office/office.asyncresult).<br/><br/>Para acessar os dados selecionados do método de retorno de chamada, chame `asyncResult.value.data`. Para acessar a propriedade original de onde a seleção é proveniente, chame `asyncResult.value.sourceProperty`, que será `body` ou `subject`.|

##### <a name="requirements"></a>Requisitos

|Requisito| Valor|
|---|---|
|[Versão mínima do conjunto de requisitos de caixa de correio](/javascript/office/requirement-sets/outlook-api-requirement-sets)| 1.2|
|[Nível de permissão mínimo](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| ReadWriteItem|
|[Modo aplicável do Outlook](https://docs.microsoft.com/outlook/add-ins/#extension-points)| Redigir|

##### <a name="returns"></a>Retorna:

Os dados selecionados em forma de sequência de caracteres com formato determinado por `coercionType`.

<dl class="param-type">

<dt>Tipo</dt>

<dd>String</dd>

</dl>

##### <a name="example"></a>Exemplo

```
// getting selected data
Office.initialize = function () {
    Office.context.mailbox.item.getSelectedDataAsync(Office.CoercionType.Text, {}, getCallback);
}

function getCallback(asyncResult) {
    var text = asyncResult.value.data;
    var prop = asyncResult.value.sourceProperty;

    Office.context.mailbox.item.setSelectedDataAsync('Setting ' + prop + ': ' + text, {}, setCallback);
}

function setCallback(asyncResult) {
    // check for errors
}
```

####  <a name="loadcustompropertiesasynccallback-usercontext"></a>loadCustomPropertiesAsync(callback, [userContext])

Carrega de forma assíncrona as propriedades personalizadas para esse suplemento no item selecionado.

Propriedades personalizadas são armazenadas como pares chave/valor de acordo com o aplicativo e o item. Este método retornará um objeto `CustomProperties` no retorno de chamada, que fornece métodos para acessar as propriedades personalizadas específicas para o item atual e o suplemento atual. As propriedades personalizadas não são criptografadas no item, portanto, isto não deve ser usado como armazenamento seguro.

##### <a name="parameters"></a>Parâmetros:

|Nome| Tipo| Atributos| Descrição|
|---|---|---|---|
|`callback`| function||Quando o método for concluído, a função passada no parâmetro `callback` é chamada com um parâmetro único, `asyncResult`, que é um objeto [`AsyncResult`](/javascript/api/office/office.asyncresult).<br/><br/>As propriedades personalizadas são fornecidas como um objeto [`CustomProperties`](/javascript/api/outlook_1_5/office.customproperties) na propriedade `asyncResult.value`. Esse objeto pode ser usado para obter, definir e remover propriedades personalizadas do item e salvar as alterações no conjunto de propriedades personalizadas de volta no servidor.|
|`userContext`| Object| &lt;opcional&gt;|Os desenvolvedores podem fornecer qualquer objeto que desejarem acessar na função de retorno de chamada. Esse objeto pode ser acessado pela propriedade `asyncResult.asyncContext` na função de retorno de chamada.|

##### <a name="requirements"></a>Requisitos

|Requisito| Valor|
|---|---|
|[Versão mínima do conjunto de requisitos de caixa de correio](/javascript/office/requirement-sets/outlook-api-requirement-sets)| 1.0|
|[Nível de permissão mínimo](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| ReadItem|
|[Modo aplicável do Outlook](https://docs.microsoft.com/outlook/add-ins/#extension-points)| Redigir ou ler|

##### <a name="example"></a>Exemplo

O exemplo de código a seguir mostra como usar o método `loadCustomPropertiesAsync` para carregar de forma assíncrona as propriedades personalizadas que são específicas para o item atual. O exemplo também mostra como usar o método `CustomProperties.saveAsync` para salvar essas propriedades de volta ao servidor. Depois de carregar as propriedades personalizadas, o exemplo de código usará o método `CustomProperties.get` para ler a propriedade personalizada `myProp`, o método `CustomProperties.set` para gravar a propriedade personalizada `otherProp` e chamará o método `saveAsync` para salvar as propriedades personalizadas.

```
// The initialize function is required for all add-ins.
Office.initialize = function () {
  // Checks for the DOM to load using the jQuery ready function.
  $(document).ready(function () {
  // After the DOM is loaded, add-in-specific code can run.
  var item = Office.context.mailbox.item;
  item.loadCustomPropertiesAsync(customPropsCallback);
  });
}

function customPropsCallback(asyncResult) {
  var customProps = asyncResult.value;
  var myProp = customProps.get("myProp");

  customProps.set("otherProp", "value");
  customProps.saveAsync(saveCallback);
}

function saveCallback(asyncResult) {
}
```

####  <a name="removeattachmentasyncattachmentid-options-callback"></a>removeAttachmentAsync(attachmentId, [options], [callback])

Remove um anexo de uma mensagem ou de um compromisso.

O método `removeAttachmentAsync` remove o anexo com o identificador especificado do item. Como prática recomendada, deve-se usar o identificador do anexo para remover um anexo somente se o mesmo aplicativo de email tiver adicionado esse anexo na mesma sessão. No Outlook Web App e no OWA para Dispositivos, o identificador do anexo é válido apenas dentro da mesma sessão. Uma sessão é finalizada quando o usuário fecha o aplicativo ou se o usuário começa a compor em um formulário embutido e, subsequentemente, sai do formulário embutido para continuar em uma janela separada.

##### <a name="parameters"></a>Parâmetros:

|Nome| Tipo| Attributes| Descrição|
|---|---|---|---|
|`attachmentId`| String||O identificador do anexo a remover. O comprimento máximo da sequência de caracteres é de 100 caracteres.|
|`options`| Object| &lt;opcional&gt;|Um literal de objeto que contém uma ou mais das propriedades a seguir.|
|`options.asyncContext`| Object| &lt;opcional&gt;|Os desenvolvedores podem fornecer qualquer objeto que desejarem acessar no método de retorno de chamada.|
|`callback`| function| &lt;opcional&gt;|Quando o método for concluído, a função passada no parâmetro `callback` é chamada com um parâmetro único, `asyncResult`, que é um objeto [`AsyncResult`](/javascript/api/office/office.asyncresult). <br/>Se a remoção do anexo falhar, a propriedade `asyncResult.error` conterá um código de erro com o motivo da falha.|

##### <a name="errors"></a>Erros

| Código de erro | Descrição |
|------------|-------------|
| `InvalidAttachmentId` | O identificador de anexo não existe. |

##### <a name="requirements"></a>Requisitos

|Requisito| Valor|
|---|---|
|[Versão mínima do conjunto de requisitos de caixa de correio](/javascript/office/requirement-sets/outlook-api-requirement-sets)| 1.1|
|[Nível de permissão mínimo](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| ReadWriteItem|
|[Modo aplicável do Outlook](https://docs.microsoft.com/outlook/add-ins/#extension-points)| Redigir|

##### <a name="example"></a>Exemplo

O código a seguir remove um anexo com um identificador '0'.

```
Office.context.mailbox.item.removeAttachmentAsync(
  '0',
  { asyncContext : null },
  function (asyncResult)
  {
    console.log(asyncResult.status);
  }
);
```

####  <a name="saveasyncoptions-callback"></a>saveAsync([options], callback)

Salva um item de forma assíncrona.

Quando chamado, este método salva a mensagem atual como um rascunho e retorna o identificador do item por meio do método de retorno de chamada. No Outlook Web App ou no Outlook no modo online, o item é salvo no servidor. No Outlook no modo cache, o item é salvo no cache local.

> [!NOTE]
> Se o seu suplemento chamar `saveAsync` em um item no modo de redação para obter um `itemId` para usar com o EWS ou a API REST, esteja ciente de que quando o Outlook estiver no modo de cache, poderá levar algum tempo antes do item realmente ser sincronizado com o servidor. Até que o item seja sincronizado, o uso de `itemId` retornará um erro.

Como compromissos não têm um estado de rascunho, se `saveAsync` for chamado em um compromisso no modo de redação, o item será salvo como um compromisso normal no calendário do usuário. Para novos compromissos que não foram salvos antes, nenhum convite será enviado. Salvar um compromisso existente enviará uma atualização aos participantes adicionados ou removidos.

> [!NOTE]
> Os seguintes clientes possuem um comportamento diferente para o `saveAsync` nos compromissos no modo de redação:
>
> - O Outlook para Mac não suporta `saveAsync` em uma reunião no modo de redação. Chamar `saveAsync` em uma reunião no Outlook para Mac retornará um erro.
> - O Outlook na Web sempre enviará um convite ou atualização quando `saveAsync` for chamado em um compromisso no modo de redação.

##### <a name="parameters"></a>Parâmetros:

|Nome| Tipo| Attributes| Descrição|
|---|---|---|---|
|`options`| Object| &lt;opcional&gt;|Um literal de objeto que contém uma ou mais das propriedades a seguir.|
|`options.asyncContext`| Object| &lt;opcional&gt;|Os desenvolvedores podem fornecer qualquer objeto que desejarem acessar no método de retorno de chamada.|
|`callback`| function||Quando o método for concluído, a função passada no parâmetro `callback` é chamada com um parâmetro único, `asyncResult`, que é um objeto [`AsyncResult`](/javascript/api/office/office.asyncresult).<br/><br/>Em caso de sucesso, o identificador do item será fornecido na propriedade `asyncResult.value`.|

##### <a name="requirements"></a>Requisitos

|Requisito| Valor|
|---|---|
|[Versão mínima do conjunto de requisitos de caixa de correio](/javascript/office/requirement-sets/outlook-api-requirement-sets)| 1.3|
|[Nível de permissão mínimo](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| ReadWriteItem|
|[Modo aplicável do Outlook](https://docs.microsoft.com/outlook/add-ins/#extension-points)| Redigir|

##### <a name="examples"></a>Exemplos

```
Office.context.mailbox.item.saveAsync(
  function callback(result) {
    // Process the result
  });
```

A seguir apresentamos um exemplo do parâmetro `result` passado para a função de retorno de chamada. A propriedade `value` contém a ID do item.

```
{
  "value":"AAMkADI5...AAA=",
  "status":"succeeded"
}
```

####  <a name="setselecteddataasyncdata-options-callback"></a>setSelectedDataAsync(data, [options], callback)

Insere de forma assíncrona os dados no corpo ou no assunto de uma mensagem.

O método `setSelectedDataAsync` insere a sequência de caracteres especificada no local do cursor no corpo ou assunto do item ou, se o texto estiver selecionado no editor, substitui o texto selecionado. Se o cursor não estiver no campo do corpo ou assunto, um erro será retornado. Após a inserção, o cursor é colocado no final do conteúdo inserido.

##### <a name="parameters"></a>Parâmetros:

|Nome| Tipo| Attributes| Descrição|
|---|---|---|---|
|`data`| String||Os dados a serem inseridos. Os dados não devem exceder 1.000.000 de caracteres. Se forem passados mais de 1.000.000 caracteres, ocorrerá uma exceção `ArgumentOutOfRange`.|
|`options`| Object| &lt;opcional&gt;|Um literal de objeto que contém uma ou mais das propriedades a seguir.|
|`options.asyncContext`| Object| &lt;opcional&gt;|Os desenvolvedores podem fornecer qualquer objeto que desejarem acessar no método de retorno de chamada.|
|`options.coercionType`| [Office.CoercionType](office.md#coerciontype-string)| &lt;opcional&gt;|Se `text`, o estilo atual é aplicado no Outlook Web App e no Outlook. Se o campo for um editor de HTML, apenas os dados de texto são inseridos, mesmo se os dados forem HTML.<br/><br/>Se `html` e o campo for compatível com HTML (e o assunto não), o estilo atual será aplicado no Outlook Web App e o estilo padrão será aplicado no Outlook. Se o campo for um campo de texto, um erro `InvalidDataFormat` será retornado.<br/><br/>Se `coercionType` não estiver definido, o resultado depende do campo: se o campo for HTML, será usado HTML; se o campo for texto, será usado texto sem formatação.|
|`callback`| function||Quando o método for concluído, a função passada no parâmetro `callback` é chamada com um parâmetro único, `asyncResult`, que é um objeto [`AsyncResult`](/javascript/api/office/office.asyncresult). |

##### <a name="requirements"></a>Requisitos

|Requisito| Valor|
|---|---|
|[Versão mínima do conjunto de requisitos de caixa de correio](/javascript/office/requirement-sets/outlook-api-requirement-sets)| 1.2|
|[Nível de permissão mínimo](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| ReadWriteItem|
|[Modo aplicável do Outlook](https://docs.microsoft.com/outlook/add-ins/#extension-points)| Redigir|

##### <a name="example"></a>Exemplo

```
Office.context.mailbox.item.setSelectedDataAsync("Hello World!");
Office.context.mailbox.item.setSelectedDataAsync("<b>Hello World!</b>", { coercionType : "html" });
```