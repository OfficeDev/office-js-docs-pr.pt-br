---
title: Office. Context. Mailbox. Item-visualização do conjunto de requisitos
description: ''
ms.date: 02/26/2019
localization_priority: Normal
ms.openlocfilehash: 32c982631dd832af6361f68176fe2c17de88b057
ms.sourcegitcommit: f7f3d38ae4430e2218bf0abe7bb2976108de3579
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 03/01/2019
ms.locfileid: "30359300"
---
# <a name="item"></a>item

### <a name="officeofficemdcontextofficecontextmdmailboxofficecontextmailboxmditem"></a>[Office](office.md)[.context](office.context.md)[.mailbox](office.context.mailbox.md).item

O namespace `item` é usado para acessar a mensagem, a solicitação de reunião ou o compromisso selecionado no momento. Você pode determinar o tipo de `item` usando a propriedade [itemType](#itemtype-officemailboxenumsitemtypejavascriptapioutlookofficemailboxenumsitemtype).

##### <a name="requirements"></a>Requisitos

|Requisito|Valor|
|---|---|
|[Versão do conjunto de requisitos mínimos da caixa de correio](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|1.0|
|[Nível de permissão mínimo](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|Restrito|
|[Modo do Outlook aplicável](https://docs.microsoft.com/outlook/add-ins/#extension-points)|Escrever ou Ler|

##### <a name="members-and-methods"></a>Membros e métodos

| Membro | Tipo |
|--------|------|
| [attachments](#attachments-arrayattachmentdetailsjavascriptapioutlookofficeattachmentdetails) | Membro |
| [bcc](#bcc-recipientsjavascriptapioutlookofficerecipients) | Membro |
| [body](#body-bodyjavascriptapioutlookofficebody) | Membro |
| [cc](#cc-arrayemailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsrecipientsjavascriptapioutlookofficerecipients) | Membro |
| [conversationId](#nullable-conversationid-string) | Membro |
| [dateTimeCreated](#datetimecreated-date) | Membro |
| [dateTimeModified](#datetimemodified-date) | Membro |
| [end](#end-datetimejavascriptapioutlookofficetime) | Membro |
| [enhancedLocation](#enhancedlocation-enhancedlocationjavascriptapioutlookofficeenhancedlocation) | Membro |
| [from](#from-emailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsfromjavascriptapioutlookofficefrom) | Membro |
| [internetHeaders](#internetheaders-internetheadersjavascriptapioutlookofficeinternetheaders) | Membro |
| [internetMessageId](#internetmessageid-string) | Membro |
| [itemClass](#itemclass-string) | Membro |
| [itemId](#nullable-itemid-string) | Membro |
| [itemType](#itemtype-officemailboxenumsitemtypejavascriptapioutlookofficemailboxenumsitemtype) | Membro |
| [location](#location-stringlocationjavascriptapioutlookofficelocation) | Membro |
| [normalizedSubject](#normalizedsubject-string) | Membro |
| [notificationMessages](#notificationmessages-notificationmessagesjavascriptapioutlookofficenotificationmessages) | Membro |
| [optionalAttendees](#optionalattendees-arrayemailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsrecipientsjavascriptapioutlookofficerecipients) | Member |
| [organizer](#organizer-emailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsorganizerjavascriptapioutlookofficeorganizer) | Membro |
| [recurrence](#nullable-recurrence-recurrencejavascriptapioutlookofficerecurrence) | Membro |
| [requiredAttendees](#requiredattendees-arrayemailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsrecipientsjavascriptapioutlookofficerecipients) | Membro |
| [sender](#sender-emailaddressdetailsjavascriptapioutlookofficeemailaddressdetails) | Membro |
| [seriesId](#nullable-seriesid-string) | Member |
| [start](#start-datetimejavascriptapioutlookofficetime) | Membro |
| [subject](#subject-stringsubjectjavascriptapioutlookofficesubject) | Membro |
| [to](#to-arrayemailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsrecipientsjavascriptapioutlookofficerecipients) | Membro |
| [addFileAttachmentAsync](#addfileattachmentasyncuri-attachmentname-options-callback) | Método |
| [addFileAttachmentFromBase64Async](#addfileattachmentfrombase64asyncbase64file-attachmentname-options-callback) | Método |
| [addHandlerAsync](#addhandlerasynceventtype-handler-options-callback) | Método |
| [addItemAttachmentAsync](#additemattachmentasyncitemid-attachmentname-options-callback) | Método |
| [close](#close) | Método |
| [displayReplyAllForm](#displayreplyallformformdata-callback) | Método |
| [displayReplyForm](#displayreplyformformdata-callback) | Método |
| [getAttachmentContentAsync](#getattachmentcontentasyncattachmentid-options-callback--attachmentcontentjavascriptapioutlookofficeattachmentcontent) | Método |
| [getAttachmentsAsync](#getattachmentsasyncoptions-callback--arrayattachmentdetailsjavascriptapioutlookofficeattachmentdetails) | Método |
| [getEntities](#getentities--entitiesjavascriptapioutlookofficeentities) | Método |
| [getEntitiesByType](#getentitiesbytypeentitytype--nullable-arraystringcontactjavascriptapioutlookofficecontactmeetingsuggestionjavascriptapioutlookofficemeetingsuggestionphonenumberjavascriptapioutlookofficephonenumbertasksuggestionjavascriptapioutlookofficetasksuggestion) | Método |
| [getFilteredEntitiesByName](#getfilteredentitiesbynamename--nullable-arraystringcontactjavascriptapioutlookofficecontactmeetingsuggestionjavascriptapioutlookofficemeetingsuggestionphonenumberjavascriptapioutlookofficephonenumbertasksuggestionjavascriptapioutlookofficetasksuggestion) | Método |
| [getInitializationContextAsync](#getinitializationcontextasyncoptions-callback) | Método |
| [getRegExMatches](#getregexmatches--object) | Método |
| [getRegExMatchesByName](#getregexmatchesbynamename--nullable-array-string-) | Método |
| [getSelectedDataAsync](#getselecteddataasynccoerciontype-options-callback--string) | Método |
| [getSelectedEntities](#getselectedentities--entitiesjavascriptapioutlookofficeentities) | Método |
| [getSelectedRegExMatches](#getselectedregexmatches--object) | Método |
| [getSharedPropertiesAsync](#getsharedpropertiesasyncoptions-callback) | Método |
| [loadCustomPropertiesAsync](#loadcustompropertiesasynccallback-usercontext) | Método |
| [removeAttachmentAsync](#removeattachmentasyncattachmentid-options-callback) | Método |
| [removeHandlerAsync](#removehandlerasynceventtype-options-callback) | Método |
| [saveAsync](#saveasyncoptions-callback) | Método |
| [setSelectedDataAsync](#setselecteddataasyncdata-options-callback) | Método |

### <a name="example"></a>Exemplo

O exemplo de código JavaScript a seguir mostra como acessar a propriedade `subject` do item atual no Outlook.

```javascript
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
};
```

### <a name="members"></a>Membros

#### <a name="attachments-arrayattachmentdetailsjavascriptapioutlookofficeattachmentdetails"></a>attachments :Array.<[AttachmentDetails](/javascript/api/outlook/office.attachmentdetails)>

Obtém os anexos do item como uma matriz. Somente modo de leitura.

> [!NOTE]
> Certos tipos de arquivos são bloqueados pelo Outlook devido a possíveis problemas de segurança e, portanto, não retornam. Para saber mais, confira [Anexos bloqueados no Outlook](https://support.office.com/article/Blocked-attachments-in-Outlook-434752E1-02D3-4E90-9124-8B81E49A8519).

##### <a name="type"></a>Tipo

*   Array.<[AttachmentDetails](/javascript/api/outlook/office.attachmentdetails)>

##### <a name="requirements"></a>Requisitos

|Requisito|Valor|
|---|---|
|[Versão do conjunto de requisitos mínimos da caixa de correio](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|1.0|
|[Nível de permissão mínimo](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|ReadItem|
|[Modo do Outlook aplicável](https://docs.microsoft.com/outlook/add-ins/#extension-points)|Read|

##### <a name="example"></a>Exemplo

O código a seguir cria uma cadeia de caracteres HTML com detalhes de todos os anexos no item atual.

```javascript
var item = Office.context.mailbox.item;
var outputString = "";

if (item.attachments.length > 0) {
  for (i = 0 ; i < item.attachments.length ; i++) {
    var attachment = item.attachments[i];
    outputString += "<BR>" + i + ". Name: ";
    outputString += attachment.name;
    outputString += "<BR>ID: " + attachment.id;
    outputString += "<BR>contentType: " + attachment.contentType;
    outputString += "<BR>size: " + attachment.size;
    outputString += "<BR>attachmentType: " + attachment.attachmentType;
    outputString += "<BR>isInline: " + attachment.isInline;
  }
}

console.log(outputString);
```

####  <a name="bcc-recipientsjavascriptapioutlookofficerecipients"></a>bcc :[Recipients](/javascript/api/outlook/office.recipients)

Obtém um objeto que fornece métodos para acessar ou atualizar os destinatários na linha Cco (com cópia oculta) de uma mensagem. Somente modo de redação.

##### <a name="type"></a>Tipo

*   [Destinatários](/javascript/api/outlook/office.recipients)

##### <a name="requirements"></a>Requisitos

|Requisito|Valor|
|---|---|
|[Versão do conjunto de requisitos mínimos da caixa de correio](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|1.1|
|[Nível de permissão mínimo](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|ReadItem|
|[Modo do Outlook aplicável](https://docs.microsoft.com/outlook/add-ins/#extension-points)|Escrever|

##### <a name="example"></a>Exemplo

```javascript
Office.context.mailbox.item.bcc.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.bcc.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.bcc.getAsync(callback);

function callback(asyncResult) {
  var arrayOfBccRecipients = asyncResult.value;
}
```

####  <a name="body-bodyjavascriptapioutlookofficebody"></a>body :[Body](/javascript/api/outlook/office.body)

Obtém um objeto que fornece métodos para manipular o corpo de um item.

##### <a name="type"></a>Tipo

*   [Body](/javascript/api/outlook/office.body)

##### <a name="requirements"></a>Requisitos

|Requisito|Valor|
|---|---|
|[Versão do conjunto de requisitos mínimos da caixa de correio](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|1.1|
|[Nível de permissão mínimo](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|ReadItem|
|[Modo do Outlook aplicável](https://docs.microsoft.com/outlook/add-ins/#extension-points)|Escrever ou Ler|

##### <a name="example"></a>Exemplo

Este exemplo obtém o corpo da mensagem em texto sem formatação.

```javascript
Office.context.mailbox.item.body.getAsync(
  "text",
  { asyncContext: "This is passed to the callback" },
  function callback(result) {
    // Do something with the result.
  });

```

A seguir apresentamos um exemplo do resultado do parâmetro passado à função de retorno de chamada.

```json
{
  "value": "TEXT of whole body (including threads below)",
  "status": "succeeded",
  "asyncContext": "This is passed to the callback"
}
```

####  <a name="cc-arrayemailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsrecipientsjavascriptapioutlookofficerecipients"></a>cc :Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook/office.recipients)

Fornece acesso aos destinatários na linha Cc (com cópia) de uma mensagem. O tipo de objeto e o nível de acesso dependem do modo do item atual.

##### <a name="read-mode"></a>Modo de leitura

A propriedade `cc` retorna uma matriz que contém um objeto `EmailAddressDetails` para cada destinatário listado na linha **Cc** da mensagem. O conjunto está limitado a um máximo de 100 membros.

```javascript
console.log(JSON.stringify(Office.context.mailbox.item.cc));
```

##### <a name="compose-mode"></a>Modo de redação

A propriedade `cc` retorna um objeto `Recipients` que fornece métodos para obter ou atualizar os destinatários na linha **Cc** da mensagem.

```javascript
Office.context.mailbox.item.cc.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.cc.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.cc.getAsync(callback);

function callback(asyncResult) {
  var arrayOfCcRecipients = asyncResult.value;
}
```

##### <a name="type"></a>Tipo

*   Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook/office.recipients)

##### <a name="requirements"></a>Requisitos

|Requisito|Valor|
|---|---|
|[Versão do conjunto de requisitos mínimos da caixa de correio](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|1.0|
|[Nível de permissão mínimo](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|ReadItem|
|[Modo do Outlook aplicável](https://docs.microsoft.com/outlook/add-ins/#extension-points)|Escrever ou Ler|

####  <a name="nullable-conversationid-string"></a>(anulável) conversationId :Cadeia de Caracteres

Obtém um identificador da conversa de email que contém uma mensagem específica.

Você pode obter um número inteiro para esta propriedade se o aplicativo de email estiver ativado nos formulários de leitura ou nas respostas em formulários de composição. Se, posteriormente, o usuário alterar o assunto da mensagem de resposta, ao enviar a resposta, a ID da conversa daquela mensagem será alterada e o valor obtido anteriormente não mais se aplicará.

Você obtém nulo para esta propriedade para um novo item em um formulário de composição. Se o usuário definir um assunto e salvar o item, a propriedade `conversationId` retornará um valor.

##### <a name="type"></a>Tipo

*   String

##### <a name="requirements"></a>Requisitos

|Requisito|Valor|
|---|---|
|[Versão do conjunto de requisitos mínimos da caixa de correio](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|1.0|
|[Nível de permissão mínimo](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|ReadItem|
|[Modo do Outlook aplicável](https://docs.microsoft.com/outlook/add-ins/#extension-points)|Escrever ou Ler|

##### <a name="example"></a>Exemplo

```javascript
var conversationId = Office.context.mailbox.item.conversationId;
console.log("conversationId: " + conversationId);
```

#### <a name="datetimecreated-date"></a>dateTimeCreated :Data

Obtém a data e a hora em que um item foi criado. Modo somente leitura.

##### <a name="type"></a>Tipo

*   Data

##### <a name="requirements"></a>Requisitos

|Requisito|Valor|
|---|---|
|[Versão do conjunto de requisitos mínimos da caixa de correio](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|1.0|
|[Nível de permissão mínimo](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|ReadItem|
|[Modo do Outlook aplicável](https://docs.microsoft.com/outlook/add-ins/#extension-points)|Read|

##### <a name="example"></a>Exemplo

```javascript
var dateTimeCreated = Office.context.mailbox.item.dateTimeCreated;
console.log("Date and time created: " + dateTimeCreated);
```

#### <a name="datetimemodified-date"></a>dateTimeModified :Date

Obtém a data e a hora em que um item foi alterado pela última vez. Somente modo de leitura.

> [!NOTE]
> Não há suporte para esse membro no Outlook para iOS ou no Outlook para Android.

##### <a name="type"></a>Tipo

*   Data

##### <a name="requirements"></a>Requisitos

|Requisito|Valor|
|---|---|
|[Versão do conjunto de requisitos mínimos da caixa de correio](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|1.0|
|[Nível de permissão mínimo](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|ReadItem|
|[Modo do Outlook aplicável](https://docs.microsoft.com/outlook/add-ins/#extension-points)|Read|

##### <a name="example"></a>Exemplo

```javascript
var dateTimeModified = Office.context.mailbox.item.dateTimeModified;
console.log("Date and time modified: " + dateTimeModified);
```

####  <a name="end-datetimejavascriptapioutlookofficetime"></a>end :Date|[Time](/javascript/api/outlook/office.time)

Obtém ou define a data e a hora em que o compromisso deve terminar.

A propriedade `end` é expressa como um valor de data e hora no Tempo Universal Coordenado (UTC). Você pode usar o método [`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttimejavascriptapioutlookofficelocalclienttime) para converter o valor da propriedade end para a data e a hora local do cliente.

##### <a name="read-mode"></a>Modo de leitura

A propriedade `end` retorna um objeto `Date`.

```javascript
var end = Office.context.mailbox.item.end;
console.log("Appointment end: " + end);
```

##### <a name="compose-mode"></a>Modo de redação

A propriedade `end` retorna um objeto `Time`.

Ao usar o método [`Time.setAsync`](/javascript/api/outlook/office.time#setasync-datetime--options--callback-) para definir a hora de término, deve-se usar o método [`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) para converter a hora local no cliente para UTC para o servidor.

O exemplo a seguir define a hora de término de um compromisso usando o método [`setAsync`](/javascript/api/outlook/office.time#setasync-datetime--options--callback-) do objeto `Time`.

```javascript
var endTime = new Date("3/14/2015");
var options = {
  // Pass information that can be used in the callback.
  asyncContext: {verb: "Set"}
};
Office.context.mailbox.item.end.setAsync(endTime, options, function(result) {
  if (result.error) {
    console.debug(result.error);
  } else {
    // Access the asyncContext that was passed to the setAsync function.
    console.debug("End Time " + result.asyncContext.verb);
  }
});
```

##### <a name="type"></a>Tipo

*   Data | [Hora](/javascript/api/outlook/office.time)

##### <a name="requirements"></a>Requisitos

|Requisito|Valor|
|---|---|
|[Versão do conjunto de requisitos mínimos da caixa de correio](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|1.0|
|[Nível de permissão mínimo](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|ReadItem|
|[Modo do Outlook aplicável](https://docs.microsoft.com/outlook/add-ins/#extension-points)|Escrever ou Ler|

#### <a name="enhancedlocation-enhancedlocationjavascriptapioutlookofficeenhancedlocation"></a>enhancedLocation:[enhancedLocation](/javascript/api/outlook/office.enhancedlocation)

Obtém ou define os locais de um compromisso.

##### <a name="read-mode"></a>Modo de leitura

A `enhancedLocation` propriedade retorna um objeto [EnhancedLocation](/javascript/api/outlook/office.enhancedlocation) que permite que você obtenha o conjunto de locais (cada um representado por um objeto [LocationDetails](/javascript/api/outlook/office.locationdetails) ) associado ao compromisso.

##### <a name="compose-mode"></a>Modo de redação

A `enhancedLocation` propriedade retorna um objeto [EnhancedLocation](/javascript/api/outlook/office.enhancedlocation) que fornece métodos para obter, remover ou adicionar locais em um compromisso.

##### <a name="type"></a>Tipo

*   [EnhancedLocation](/javascript/api/outlook/office.enhancedlocation)

##### <a name="requirements"></a>Requisitos

|Requisito|Valor|
|---|---|
|[Versão do conjunto de requisitos mínimos da caixa de correio](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|Visualização|
|[Nível de permissão mínimo](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|ReadItem|
|[Modo do Outlook aplicável](https://docs.microsoft.com/outlook/add-ins/#extension-points)|Escrever ou Ler|

##### <a name="example"></a>Exemplo

O exemplo a seguir obtém os locais atuais associados ao compromisso.

```javascript
Office.context.mailbox.item.enhancedLocation.getAsync(callbackFunction);

function callbackFunction(asyncResult) {
  asyncResult.value.forEach(function (place) {
    console.log("Display name: " + place.displayName);
    console.log("Type: " + place.locationIdentifier.type);
    if (place.locationIdentifier.type === Office.MailboxEnums.LocationType.Room) {
      console.log("Email address: " + place.emailAddress);
    }
  });
}
```

#### <a name="from-emailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsfromjavascriptapioutlookofficefrom"></a>from :[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)|[From](/javascript/api/outlook/office.from)

Obtém o endereço de email do remetente de uma mensagem.

As propriedades `from` e [`sender`](#sender-emailaddressdetailsjavascriptapioutlookofficeemailaddressdetails) representam a mesma pessoa, a menos que a mensagem seja enviada por um representante. Nesse caso, a propriedade `from` representa o delegante, e a propriedade sender, o representante.

> [!NOTE]
> A propriedade `recipientType` do objeto `EmailAddressDetails` na propriedade `from` é `undefined`.

##### <a name="read-mode"></a>Modo de leitura

A propriedade `from` retorna um objeto `EmailAddressDetails`.

```javascript
var from = Office.context.mailbox.item.from;
console.log("From " + from);
```

##### <a name="compose-mode"></a>Modo de redação

A propriedade `from` retorna um objeto `From` que fornece um método para obtenção do valor de from.

```javascript
Office.context.mailbox.item.from.getAsync(callback);

function callback(asyncResult) {
  var from = asyncResult.value;
}
```

##### <a name="type"></a>Tipo

*   [EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails) | [From](/javascript/api/outlook/office.from)

##### <a name="requirements"></a>Requisitos

|Requisito|||
|---|---|---|
|[Versão do conjunto de requisitos mínimos da caixa de correio](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|1.0|1.7|
|[Nível de permissão mínimo](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|ReadItem|ReadWriteItem|
|[Modo do Outlook aplicável](https://docs.microsoft.com/outlook/add-ins/#extension-points)|Read|Escrever|

#### <a name="internetheaders-internetheadersjavascriptapioutlookofficeinternetheaders"></a>internetHeaders :[InternetHeaders](/javascript/api/outlook/office.internetheaders)

Obtém ou define os cabeçalhos de internet de uma mensagem.

##### <a name="type"></a>Tipo

*   [InternetHeaders](/javascript/api/outlook/office.internetheaders)

##### <a name="requirements"></a>Requisitos

|Requisito|Valor|
|---|---|
|[Versão do conjunto de requisitos mínimos da caixa de correio](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|Visualização|
|[Nível de permissão mínimo](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|ReadItem|
|[Modo do Outlook aplicável](https://docs.microsoft.com/outlook/add-ins/#extension-points)|Escrever ou Ler|

##### <a name="example"></a>Exemplo

```javascript
Office.context.mailbox.item.internetHeaders.getAsync(["header1", "header2"], callback);

function callback(asyncResult) {
  var dictionary = asyncResult.value;
  var header1_value = dictionary["header1"];
}
```

#### <a name="internetmessageid-string"></a>internetMessageId Cadeia de Caracteres

Obtém o identificador de mensagem de Internet para uma mensagem de email. Modo somente leitura.

##### <a name="type"></a>Tipo

*   String

##### <a name="requirements"></a>Requisitos

|Requisito|Valor|
|---|---|
|[Versão do conjunto de requisitos mínimos da caixa de correio](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|1.0|
|[Nível de permissão mínimo](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|ReadItem|
|[Modo do Outlook aplicável](https://docs.microsoft.com/outlook/add-ins/#extension-points)|Leitura|

##### <a name="example"></a>Exemplo

```javascript
var internetMessageId = Office.context.mailbox.item.internetMessageId;
console.log("internetMessageId: " + internetMessageId);
```

#### <a name="itemclass-string"></a>itemClass :String

Obtém a classe do item dos Serviços Web do Exchange do item selecionado. Somente modo de leitura.

A propriedade `itemClass` especifica a classe da mensagem do item selecionado. A seguir estão as classes de mensagem padrão para o item de mensagem ou de compromisso.

|Tipo	|Descrição|classe de item|
|---|---|---|
|Itens de compromisso|Esses são itens de calendário da classe de item `IPM.Appointment` ou `IPM.Appointment.Occurrence`.|`IPM.Appointment`<br />`IPM.Appointment.Occurrence`|
|Itens de mensagem|Incluem mensagens de email que têm a classe de mensagem padrão `IPM.Note` e solicitações de reunião, respostas e cancelamentos, que utilizam `IPM.Schedule.Meeting` como a classe de mensagem básica.|`IPM.Note`<br />`IPM.Schedule.Meeting.Request`<br />`IPM.Schedule.Meeting.Neg`<br />`IPM.Schedule.Meeting.Pos`<br />`IPM.Schedule.Meeting.Tent`<br />`IPM.Schedule.Meeting.Canceled`|

Você pode criar classes de mensagem personalizadas que estendem uma classe de mensagem padrão, por exemplo, uma classe de mensagem de compromisso `IPM.Appointment.Contoso` personalizada.

##### <a name="type"></a>Tipo

*   String

##### <a name="requirements"></a>Requisitos

|Requisito|Valor|
|---|---|
|[Versão do conjunto de requisitos mínimos da caixa de correio](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|1.0|
|[Nível de permissão mínimo](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|ReadItem|
|[Modo do Outlook aplicável](https://docs.microsoft.com/outlook/add-ins/#extension-points)|Leitura|

##### <a name="example"></a>Exemplo

```javascript
var itemClass = Office.context.mailbox.item.itemClass;
console.log("Item class: " + itemClass);
```

#### <a name="nullable-itemid-string"></a>(nullable) itemId :String

Obtém o identificador do item dos Serviços Web do Exchange para o item atual. Somente modo de leitura.

> [!NOTE]
> O identificador retornado pela propriedade `itemId` é o mesmo que o identificador do item dos Serviços Web do Exchange. A propriedade `itemId` não é idêntica à ID de Entrada do Outlook ou a ID usada pela API REST do Outlook. Antes de fazer chamadas API REST usando esse valor, ela deverá ser convertida usando [Office.context.mailbox.convertToRestId](office.context.mailbox.md#converttorestiditemid-restversion--string). Para obter detalhes, confira [Usar APIs REST do Outlook de um suplemento do Outlook](https://docs.microsoft.com/outlook/add-ins/use-rest-api#get-the-item-id).

A propriedade `itemId` não está disponível no modo de redação. Se for obrigatório o identificador de um item, pode ser usado o método [`saveAsync`](#saveasyncoptions-callback) para salvar o item no servidor, o que retornará o identificador do item no parâmetro [`AsyncResult.value`](/javascript/api/office/office.asyncresult) na função de retorno de chamada.

##### <a name="type"></a>Tipo

*   String

##### <a name="requirements"></a>Requisitos

|Requisito|Valor|
|---|---|
|[Versão do conjunto de requisitos mínimos da caixa de correio](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|1.0|
|[Nível de permissão mínimo](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|ReadItem|
|[Modo do Outlook aplicável](https://docs.microsoft.com/outlook/add-ins/#extension-points)|Leitura|

##### <a name="example"></a>Exemplo

O código a seguir verifica a presença de um identificador de item. Se a propriedade `itemId` retorna `null` ou `undefined`, ele salva o item no repositório e obtém o identificador do item do resultado assíncrono.

```javascript
var itemId = Office.context.mailbox.item.itemId;
if (itemId === null || itemId == undefined) {
  Office.context.mailbox.item.saveAsync(function(result) {
    itemId = result.value;
  });
}
```

####  <a name="itemtype-officemailboxenumsitemtypejavascriptapioutlookofficemailboxenumsitemtype"></a>itemType :[Office.MailboxEnums.ItemType](/javascript/api/outlook/office.mailboxenums.itemtype)

Obtém o tipo de item que representa uma instância.

A propriedade `itemType` retorna um dos valores de enumeração `ItemType`, indicando se a instância do objeto `item` é uma mensagem ou um compromisso.

##### <a name="type"></a>Tipo

*   [Office.MailboxEnums.ItemType](/javascript/api/outlook/office.mailboxenums.itemtype)

##### <a name="requirements"></a>Requisitos

|Requisito|Valor|
|---|---|
|[Versão do conjunto de requisitos mínimos da caixa de correio](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|1.0|
|[Nível de permissão mínimo](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|ReadItem|
|[Modo do Outlook aplicável](https://docs.microsoft.com/outlook/add-ins/#extension-points)|Escrever ou Ler|

##### <a name="example"></a>Exemplo

```javascript
if (Office.context.mailbox.item.itemType === Office.MailboxEnums.ItemType.Message) {
  // Do something.
} else {
  // Do something else.
}
```

####  <a name="location-stringlocationjavascriptapioutlookofficelocation"></a>location :String|[Location](/javascript/api/outlook/office.location)

Obtém ou define o local de um compromisso.

##### <a name="read-mode"></a>Modo de leitura

A propriedade `location` retorna uma cadeia de caracteres que contém o local do compromisso.

```javascript
var location = Office.context.mailbox.item.location;
console.log("location: " + location);
```

##### <a name="compose-mode"></a>Modo de redação

A propriedade `location` retorna um objeto `Location` que fornece os métodos usados para obter e definir o local do compromisso.

```javascript
var userContext = { value : 1 };
Office.context.mailbox.item.location.getAsync( { context: userContext}, callback);

function callback(asyncResult) {
  var context = asyncResult.context;
  var location = asyncResult.value;
}
```

##### <a name="type"></a>Tipo

*   Cadeia de caracteres | [Localização](/javascript/api/outlook/office.location)

##### <a name="requirements"></a>Requisitos

|Requisito|Valor|
|---|---|
|[Versão do conjunto de requisitos mínimos da caixa de correio](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|1.0|
|[Nível de permissão mínimo](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|ReadItem|
|[Modo do Outlook aplicável](https://docs.microsoft.com/outlook/add-ins/#extension-points)|Escrever ou Ler|

#### <a name="normalizedsubject-string"></a>normalizedSubject :Cadeia de caracteres

Obtém o assunto de um item, com todos os prefixos removidos (incluindo `RE:` e `FWD:`). Somente modo de leitura.

A propriedade normalizedSubject obtém o assunto do item, com quaisquer prefixos padrão (como `RE:` e `FW:`), que são adicionados por programas de email. Para obter o assunto do item com os prefixos intactos, use a propriedade [`subject`](#subject-stringsubjectjavascriptapioutlookofficesubject).

##### <a name="type"></a>Tipo

*   String

##### <a name="requirements"></a>Requisitos

|Requisito|Valor|
|---|---|
|[Versão do conjunto de requisitos mínimos da caixa de correio](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|1.0|
|[Nível de permissão mínimo](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|ReadItem|
|[Modo do Outlook aplicável](https://docs.microsoft.com/outlook/add-ins/#extension-points)|Leitura|

##### <a name="example"></a>Exemplo

```javascript
var normalizedSubject = Office.context.mailbox.item.normalizedSubject;
console.log("Normalized subject: " + normalizedSubject);
```

####  <a name="notificationmessages-notificationmessagesjavascriptapioutlookofficenotificationmessages"></a>notificationMessages :[NotificationMessages](/javascript/api/outlook/office.notificationmessages)

Obtém as mensagens de notificação de um item.

##### <a name="type"></a>Tipo

*   [NotificationMessages](/javascript/api/outlook/office.notificationmessages)

##### <a name="requirements"></a>Requisitos

|Requisito|Valor|
|---|---|
|[Versão do conjunto de requisitos mínimos da caixa de correio](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|1.3|
|[Nível de permissão mínimo](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|ReadItem|
|[Modo do Outlook aplicável](https://docs.microsoft.com/outlook/add-ins/#extension-points)|Escrever ou Ler|

##### <a name="example"></a>Exemplo

```javascript
// Get all notifications.
Office.context.mailbox.item.notificationMessages.getAllAsync(
  function (asyncResult) {
    console.log(JSON.stringify(asyncResult));
  }
);
```

####  <a name="optionalattendees-arrayemailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsrecipientsjavascriptapioutlookofficerecipients"></a>optionalAttendees :Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook/office.recipients)

Fornece acesso aos participantes opcionais de um evento. O tipo de objeto e o nível de acesso dependem do modo do item atual.

##### <a name="read-mode"></a>Modo de leitura

A propriedade `optionalAttendees` retorna uma matriz que contém um objeto `EmailAddressDetails` para cada participante opcional da reunião.

```javascript
var optionalAttendees = Office.context.mailbox.item.optionalAttendees;
console.log("Optional attendees: " + JSON.stringify(optionalAttendees));
```

##### <a name="compose-mode"></a>Modo de redação

A propriedade `optionalAttendees` retorna um objeto `Recipients` que fornece métodos para obter ou atualizar os participantes opcionais de uma reunião.

```javascript
Office.context.mailbox.item.optionalAttendees.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.optionalAttendees.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.optionalAttendees.getAsync(callback);

function callback(asyncResult) {
  var arrayOfOptionalAttendeesRecipients = asyncResult.value;
}
```

##### <a name="type"></a>Tipo

*   Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook/office.recipients)

##### <a name="requirements"></a>Requisitos

|Requisito|Valor|
|---|---|
|[Versão do conjunto de requisitos mínimos da caixa de correio](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|1.0|
|[Nível de permissão mínimo](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|ReadItem|
|[Modo do Outlook aplicável](https://docs.microsoft.com/outlook/add-ins/#extension-points)|Escrever ou Ler|

#### <a name="organizer-emailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsorganizerjavascriptapioutlookofficeorganizer"></a>organizer :[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)|[Organizer](/javascript/api/outlook/office.organizer)

Obtém o endereço de email do organizador para uma reunião especificada.

##### <a name="read-mode"></a>Modo de leitura

A propriedade `organizer` retorna um objeto [EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails) que representa o organizador da reunião.

```javascript
var organizerName = Office.context.mailbox.item.organizer.displayName;
var organizerAddress = Office.context.mailbox.item.organizer.emailAddress;
console.log("Organizer: " + organizerName + " (" + organizerAddress + ")");
```

##### <a name="compose-mode"></a>Modo de redação

A propriedade `organizer` retorna um objeto [Organizer](/javascript/api/outlook/office.organizer) que fornece um método para obtenção do valor de organizer.

```javascript
Office.context.mailbox.item.organizer.getAsync(
  function(asyncResult) {
    console.log(JSON.stringify(asyncResult));
  }
);
```

##### <a name="type"></a>Tipo

*   [EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails) | [Organizer](/javascript/api/outlook/office.organizer)

##### <a name="requirements"></a>Requisitos

|Requisito|||
|---|---|---|
|[Versão do conjunto de requisitos mínimos da caixa de correio](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|1.0|1.7|
|[Nível de permissão mínimo](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|ReadItem|ReadWriteItem|
|[Modo do Outlook aplicável](https://docs.microsoft.com/outlook/add-ins/#extension-points)|Read|Escrever|

#### <a name="nullable-recurrence-recurrencejavascriptapioutlookofficerecurrence"></a>(nullable) recurrence :[Recurrence](/javascript/api/outlook/office.recurrence)

Obtém ou configura o padrão de recorrência de um compromisso. Obtém o padrão de recorrência de uma solicitação de reunião. Modos de leitura e redação para itens do compromisso. Modo de leitura para os itens de solicitação de reunião.

A propriedade `recurrence` retorna um objeto [recurrence](/javascript/api/outlook/office.recurrence) para compromissos recorrentes ou solicitações de reuniões se um item for uma série ou uma instância em uma série. `null` retorna para compromissos individuais e solicitações de reunião de compromissos individuais. `undefined` retorna para mensagens que não são solicitações de reunião.

> Observação: solicitações de reunião têm um valor `itemClass` de IPM.Schedule.Meeting.Request.

> Observação: se o objeto de recorrência for `null`, isso indicará que o objeto é um compromisso individual ou uma solicitação de reunião de um compromisso individual e NÃO parte de uma série.

##### <a name="read-mode"></a>Modo de leitura

A `recurrence` propriedade retorna um objeto [Recurrence](/javascript/api/outlook/office.recurrence) que representa a recorrência do compromisso. Isso está disponível para compromissos e solicitações de reunião.

```javascript
var recurrence = Office.context.mailbox.item.recurrence;
console.log("Recurrence: " + JSON.stringify(recurrence));
```

##### <a name="compose-mode"></a>Modo de redação

A `recurrence` propriedade retorna um objeto [Recurrence](/javascript/api/outlook/office.recurrence) que fornece métodos para gerenciar a recorrência do compromisso. Isso está disponível para compromissos.

```javascript
Office.context.mailbox.item.recurrence.getAsync(callback);

function callback(asyncResult) {
  var context = asyncResult.context;
  var recurrence = asyncResult.value;
  if (!recurrence) {
    console.log("One-time appointment or meeting");
  } else {
    console.log(JSON.stringify(recurrence));
  }
}

// The following example shows the results of the getAsync call that retrieves the recurrence for a series.
// NOTE: In this example, seriesTimeObject is a placeholder for the JSON representing the
// recurrence.seriesTime property. You should use the SeriesTime object's methods to get the
// recurrence date and time properties.
Recurrence = {
  "recurrenceType": "weekly",
  "recurrenceProperties": {"interval": 2, "days": ["mon","thu","fri"], "firstDayOfWeek": "sun"},
  "seriesTime": {seriesTimeObject},
  "recurrenceTimeZone": {"name": "Pacific Standard Time", "offset": -480}
}
```

##### <a name="type"></a>Tipo

* [Recurrence](/javascript/api/outlook/office.recurrence)

|Requisito|Valor|
|---|---|
|[Versão do conjunto de requisitos mínimos da caixa de correio](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|1.7|
|[Nível de permissão mínimo](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|ReadItem|
|[Modo do Outlook aplicável](https://docs.microsoft.com/outlook/add-ins/#extension-points)|Escrever ou Ler|

####  <a name="requiredattendees-arrayemailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsrecipientsjavascriptapioutlookofficerecipients"></a>requiredAttendees :Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook/office.recipients)

Fornece acesso aos participantes obrigatórios de um evento. O tipo de objeto e o nível de acesso dependem do modo do item atual.

##### <a name="read-mode"></a>Modo de leitura

A propriedade `requiredAttendees` retorna uma matriz que contém um objeto `EmailAddressDetails` para cada participante obrigatório da reunião.

```javascript
var requiredAttendees = Office.context.mailbox.item.requiredAttendees;
console.log("Required attendees: " + JSON.stringify(requiredAttendees));
```

##### <a name="compose-mode"></a>Modo de redação

A propriedade `requiredAttendees` retorna um objeto `Recipients` que fornece métodos para obter ou atualizar os participantes obrigatórios de uma reunião.

```javascript
Office.context.mailbox.item.requiredAttendees.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.requiredAttendees.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.requiredAttendees.getAsync(callback);

function callback(asyncResult) {
  var arrayOfRequiredAttendeesRecipients = asyncResult.value;
  console.log(JSON.stringify(arrayOfRequiredAttendeesRecipients));
}
```

##### <a name="type"></a>Tipo

*   Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook/office.recipients)

##### <a name="requirements"></a>Requisitos

|Requisito|Valor|
|---|---|
|[Versão do conjunto de requisitos mínimos da caixa de correio](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|1.0|
|[Nível de permissão mínimo](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|ReadItem|
|[Modo do Outlook aplicável](https://docs.microsoft.com/outlook/add-ins/#extension-points)|Escrever ou Ler|

#### <a name="sender-emailaddressdetailsjavascriptapioutlookofficeemailaddressdetails"></a>remetente :[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)

Obtém o endereço de email do remetente de uma mensagem de email. Somente modo de leitura.

As propriedades [`from`](#from-emailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsfromjavascriptapioutlookofficefrom) e `sender` representam a mesma pessoa, a menos que a mensagem seja enviada por um representante. Nesse caso, a propriedade `from` representa o delegante, e a propriedade sender, o representante.

> [!NOTE]
> A propriedade `recipientType` do objeto `EmailAddressDetails` na propriedade `sender` é `undefined`.

##### <a name="type"></a>Tipo

*   [EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)

##### <a name="requirements"></a>Requisitos

|Requisito|Valor|
|---|---|
|[Versão do conjunto de requisitos mínimos da caixa de correio](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|1.0|
|[Nível de permissão mínimo](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|ReadItem|
|[Modo do Outlook aplicável](https://docs.microsoft.com/outlook/add-ins/#extension-points)|Read|

##### <a name="example"></a>Exemplo

```javascript
var senderName = Office.context.mailbox.item.sender.displayName;
var senderAddress = Office.context.mailbox.item.sender.emailAddress;
console.log("Sender: " + senderName + " (" + senderAddress + ")");
```

#### <a name="nullable-seriesid-string"></a>(nullable) seriesId :String

Obtém a id da série a qual uma instância pertence.

No OWA e no Outlook, o `seriesId` retorna a ID dos Serviços Web do Exchange (EWS) do item pai (série) a qual este item pertence. No entanto, no iOS e no Android, o `seriesId` retorna a ID REST do item pai.

> [!NOTE]
> O identificador retornado pela propriedade `seriesId` é o mesmo que o identificador do item dos Serviços Web do Exchange. A propriedade `seriesId` não é idêntica à ID do Outlook usada pela API REST do Outlook. Antes de fazer chamadas API REST usando esse valor, ela deverá ser convertida usando [Office.context.mailbox.convertToRestId](office.context.mailbox.md#converttorestiditemid-restversion--string). Para obter detalhes, confira [Usar APIs REST do Outlook de um suplemento do Outlook](https://docs.microsoft.com/outlook/add-ins/use-rest-api).

A propriedade `seriesId` retorna `null` para itens que não têm itens pai como compromissos individuais, itens de série ou solicitações de reunião e retorna `undefined` para outros itens que não são solicitações de reunião.

##### <a name="type"></a>Tipo

* String

##### <a name="requirements"></a>Requisitos

|Requisito|Valor|
|---|---|
|[Versão do conjunto de requisitos mínimos da caixa de correio](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|1.7|
|[Nível de permissão mínimo](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|ReadItem|
|[Modo do Outlook aplicável](https://docs.microsoft.com/outlook/add-ins/#extension-points)|Escrever ou Ler|

##### <a name="example"></a>Exemplo

```javascript
var seriesId = Office.context.mailbox.item.seriesId;

// The seriesId property returns null for items that do
// not have parent items (such as single appointments,
// series items, or meeting requests) and returns
// undefined for messages that are not meeting requests.
var isSeriesInstance = (seriesId != null);
console.log("SeriesId is " + seriesId + " and isSeriesInstance is " + isSeriesInstance);
```

####  <a name="start-datetimejavascriptapioutlookofficetime"></a>start :Date|[Time](/javascript/api/outlook/office.time)

Obtém ou define a data e a hora em que o compromisso deve começar.

A propriedade `start` é expressa como um valor de data e hora no Tempo Universal Coordenado (UTC). Você pode usar o método [`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttimejavascriptapioutlookofficelocalclienttime) para converter o valor para a data e a hora local do cliente.

##### <a name="read-mode"></a>Modo de leitura

A propriedade `start` retorna um objeto `Date`.

```javascript
var start = Office.context.mailbox.item.start;
console.log("Appointment start: " + JSON.stringify(start));
```

##### <a name="compose-mode"></a>Modo de redação

A propriedade `start` retorna um objeto `Time`.

Ao usar o método [`Time.setAsync`](/javascript/api/outlook/office.time#setasync-datetime--options--callback-) para definir a hora de início, deve-se usar o método [`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) para converter a hora local no cliente para UTC para o servidor.

O exemplo a seguir define a hora de início de um compromisso no modo de composição usando o método [`setAsync`](/javascript/api/outlook/office.time#setasync-datetime--options--callback-) do objeto `Time`.

```javascript
var startTime = new Date("3/14/2015");
var options = {
  // Pass information that can be used in the callback.
  asyncContext: {verb: "Set"}
};
Office.context.mailbox.item.start.setAsync(startTime, options, function(result) {
  if (result.error) {
    console.debug(result.error);
  } else {
    // Access the asyncContext that was passed to the setAsync function.
    console.debug("Start Time " + result.asyncContext.verb);
  }
});
```

##### <a name="type"></a>Tipo

*   Data | [Hora](/javascript/api/outlook/office.time)

##### <a name="requirements"></a>Requisitos

|Requisito|Valor|
|---|---|
|[Versão do conjunto de requisitos mínimos da caixa de correio](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|1.0|
|[Nível de permissão mínimo](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|ReadItem|
|[Modo do Outlook aplicável](https://docs.microsoft.com/outlook/add-ins/#extension-points)|Escrever ou Ler|

####  <a name="subject-stringsubjectjavascriptapioutlookofficesubject"></a>subject :String|[Subject](/javascript/api/outlook/office.subject)

Obtém ou define a descrição que aparece no campo de assunto de um item.

A propriedade `subject` obtém ou define o assunto completo do item, conforme enviado pelo servidor de email.

##### <a name="read-mode"></a>Modo de leitura

A propriedade `subject` retorna uma cadeia de caracteres. Use a propriedade [`normalizedSubject`](#normalizedsubject-string) para obter o assunto, exceto pelos prefixos iniciais, como `RE:` e `FW:`.

O exemplo de código JavaScript a seguir mostra como acessar a propriedade `subject` do item atual no Outlook.

```javascript
var subject = Office.context.mailbox.item.subject;
console.log(subject);
```

##### <a name="compose-mode"></a>Modo de redação
A propriedade `subject` retorna um objeto `Subject` que fornece métodos para obter e definir o assunto.

```javascript
Office.context.mailbox.item.subject.getAsync(callback);

function callback(asyncResult) {
  var subject = asyncResult.value;
  console.log(subject);
}
```

##### <a name="type"></a>Tipo

*   String | [Subject](/javascript/api/outlook/office.subject)

##### <a name="requirements"></a>Requisitos

|Requisito|Valor|
|---|---|
|[Versão do conjunto de requisitos mínimos da caixa de correio](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|1.0|
|[Nível de permissão mínimo](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|ReadItem|
|[Modo do Outlook aplicável](https://docs.microsoft.com/outlook/add-ins/#extension-points)|Escrever ou Ler|

####  <a name="to-arrayemailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsrecipientsjavascriptapioutlookofficerecipients"></a>to :Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook/office.recipients)

Fornece acesso aos destinatários na linha **Para** de uma mensagem. O tipo de objeto e o nível de acesso dependem do modo do item atual.

##### <a name="read-mode"></a>Modo de leitura

A propriedade `to` retorna uma matriz que contém um objeto `EmailAddressDetails` para cada destinatário listado na linha **Para** da mensagem. O conjunto está limitado a um máximo de 100 membros.

```javascript
console.log(JSON.stringify(Office.context.mailbox.item.to));
```

##### <a name="compose-mode"></a>Modo de redação

A propriedade `to` retorna um objeto `Recipients` que fornece métodos para obter ou atualizar os destinatários na linha **Para** da mensagem.

```javascript
Office.context.mailbox.item.to.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.to.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.to.getAsync(callback);

function callback(asyncResult) {
  var arrayOfToRecipients = asyncResult.value;
}
```

##### <a name="type"></a>Tipo

*   Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook/office.recipients)

##### <a name="requirements"></a>Requisitos

|Requisito|Valor|
|---|---|
|[Versão do conjunto de requisitos mínimos da caixa de correio](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|1.0|
|[Nível de permissão mínimo](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|ReadItem|
|[Modo do Outlook aplicável](https://docs.microsoft.com/outlook/add-ins/#extension-points)|Escrever ou Ler|

### <a name="methods"></a>Métodos

####  <a name="addfileattachmentasyncuri-attachmentname-options-callback"></a>addFileAttachmentAsync(uri, attachmentName, [options], [callback])

Adiciona um arquivo a uma mensagem ou um compromisso como um anexo.

O método `addFileAttachmentAsync` carrega o arquivo no URI especificado e anexa-o ao item no formulário de composição.

Posteriormente, você poderá usar o identificador com o método [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) para remover o anexo na mesma sessão.

##### <a name="parameters"></a>Parâmetros
|Nome|Tipo|Atributos|Descrição|
|---|---|---|---|
|`uri`|Cadeia de caracteres||O URI que fornece o local do arquivo anexado à mensagem ou compromisso. O comprimento máximo é de 2048 caracteres.|
|`attachmentName`|String||O nome do anexo que é mostrado enquanto o anexo está sendo carregado. O tamanho máximo é de 255 caracteres.|
|`options`|Object|&lt;opcional&gt;|Um objeto literal que contém uma ou mais das propriedades a seguir.|
|`options.asyncContext`|Objeto|&lt;opcional&gt;|Os desenvolvedores podem fornecer qualquer objeto que desejarem acessar no método de retorno de chamada.|
|`options.isInline`|Booliano|&lt;opcional&gt;|Se for `true`, indicará que o anexo será mostrado embutido no corpo da mensagem e não deverá ser exibido na lista de anexos.|
|`callback`|function|&lt;opcional&gt;|Quando o método for concluído, a função passada ao parâmetro `callback` é chamada com um único parâmetro, `asyncResult`, que é um objeto [`AsyncResult`](/javascript/api/office/office.asyncresult). <br/>Em caso de êxito, o identificador do anexo será fornecido na propriedade `asyncResult.value`.<br/>Se houver falha ao carregar o anexo, o objeto `asyncResult` conterá um objeto `Error` que fornece uma descrição do erro.|

##### <a name="errors"></a>Erros

|Código de erro|Descrição|
|------------|-------------|
|`AttachmentSizeExceeded`|O anexo é maior do que permitido.|
|`FileTypeNotSupported`|O anexo tem uma extensão que não é permitida.|
|`NumberOfAttachmentsExceeded`|A mensagem ou o compromisso tem muitos anexos.|

##### <a name="requirements"></a>Requisitos

|Requisito|Valor|
|---|---|
|[Versão do conjunto de requisitos mínimos da caixa de correio](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|1.1|
|[Nível de permissão mínimo](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|ReadWriteItem|
|[Modo do Outlook aplicável](https://docs.microsoft.com/outlook/add-ins/#extension-points)|Escrever|

##### <a name="examples"></a>Exemplos

```javascript
function callback(result) {
  if (result.error) {
    console.log(result.error);
  } else {
    console.log("Attachment added");
  }
}

function addAttachment() {
  // The values in asyncContext can be accessed in the callback.
  var options = { 'asyncContext': { var1: 1, var2: 2 } };

  var attachmentURL = "https://contoso.com/rtm/icon.png";
  Office.context.mailbox.item.addFileAttachmentAsync(attachmentURL, attachmentURL, options, callback);
}
```

O exemplo a seguir adiciona um arquivo de imagem como um anexo embutido e faz referência ao anexo no corpo da mensagem.

```javascript
Office.context.mailbox.item.addFileAttachmentAsync(
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
        // Do something here.
      });
  });
```

#### <a name="addfileattachmentfrombase64asyncbase64file-attachmentname-options-callback"></a>addFileAttachmentFromBase64Async(base64File, attachmentName, [options], [callback])

Adiciona um arquivo a partir da codificação base64 a uma mensagem ou compromisso como anexo.

O método `addFileAttachmentFromBase64Async` carrega o arquivo a partir da codificação base64 e o anexa ao item no formulário de composição. Esse método retorna o identificador de anexo no objeto AsyncResult.value.

Posteriormente, você poderá usar o identificador com o método [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) para remover o anexo na mesma sessão.

##### <a name="parameters"></a>Parâmetros
|Nome|Tipo|Atributos|Descrição|
|---|---|---|---|
|`base64File`|String||O conteúdo codificado em Base 64 de uma imagem ou arquivo a ser adicionado a um email ou evento.|
|`attachmentName`|String||O nome do anexo que é mostrado enquanto o anexo está sendo carregado. O tamanho máximo é de 255 caracteres.|
|`options`|Object|&lt;opcional&gt;|Um objeto literal que contém uma ou mais das propriedades a seguir.|
|`options.asyncContext`|Objeto|&lt;opcional&gt;|Os desenvolvedores podem fornecer qualquer objeto que desejarem acessar no método de retorno de chamada.|
|`options.isInline`|Booliano|&lt;opcional&gt;|Se for `true`, indicará que o anexo será mostrado embutido no corpo da mensagem e não deverá ser exibido na lista de anexos.|
|`callback`|function|&lt;opcional&gt;|Quando o método for concluído, a função passada ao parâmetro `callback` é chamada com um único parâmetro, `asyncResult`, que é um objeto [`AsyncResult`](/javascript/api/office/office.asyncresult). <br/>Em caso de êxito, o identificador do anexo será fornecido na propriedade `asyncResult.value`.<br/>Se houver falha ao carregar o anexo, o objeto `asyncResult` conterá um objeto `Error` que fornece uma descrição do erro.|

##### <a name="errors"></a>Erros

|Código de erro|Descrição|
|------------|-------------|
|`AttachmentSizeExceeded`|O anexo é maior do que permitido.|
|`FileTypeNotSupported`|O anexo tem uma extensão que não é permitida.|
|`NumberOfAttachmentsExceeded`|A mensagem ou o compromisso tem muitos anexos.|

##### <a name="requirements"></a>Requisitos

|Requisito|Valor|
|---|---|
|[Versão do conjunto de requisitos mínimos da caixa de correio](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|Visualização|
|[Nível de permissão mínimo](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|ReadWriteItem|
|[Modo do Outlook aplicável](https://docs.microsoft.com/outlook/add-ins/#extension-points)|Redação|

##### <a name="examples"></a>Exemplos

```javascript
Office.context.mailbox.item.addFileAttachmentFromBase64Async(
  base64String,
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
        // Do something here.
      });
  });
```

####  <a name="addhandlerasynceventtype-handler-options-callback"></a>addHandlerAsync(eventType, handler, [options], [callback])

Adiciona um manipulador de eventos a um evento com suporte.

Atualmente, os tipos de eventos `Office.EventType.AttachmentsChanged`suportados `Office.EventType.AppointmentTimeChanged`são `Office.EventType.EnhancedLocationsChanged`, `Office.EventType.RecipientsChanged`,, `Office.EventType.RecurrenceChanged`e.

##### <a name="parameters"></a>Parâmetros

| Nome | Tipo | Atributos | Descrição |
|---|---|---|---|
| `eventType` | [Office.EventType](office.md#eventtype-string) || O evento que deve invocar o manipulador. |
| `handler` | Função || A função para manipular o evento. A função deve aceitar um parâmetro exclusivo, que é um objeto literal. A propriedade `type` no parâmetro corresponderá ao parâmetro `eventType` passado para `addHandlerAsync`. |
| `options` | Objeto | &lt;opcional&gt; | Um objeto literal que contém uma ou mais das propriedades a seguir. |
| `options.asyncContext` | Objeto | &lt;opcional&gt; | Os desenvolvedores podem fornecer qualquer objeto que desejarem acessar no método de retorno de chamada. |
| `callback` | function| &lt;opcional&gt;|Quando o método for concluído, a função passada ao parâmetro `callback` é chamada com um único parâmetro, `asyncResult`, que é um objeto [`AsyncResult`](/javascript/api/office/office.asyncresult).|

##### <a name="requirements"></a>Requisitos

|Requisito| Valor|
|---|---|
|[Versão do conjunto de requisitos mínimos da caixa de correio](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| 1.7 |
|[Nível de permissão mínimo](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| ReadItem |
|[Modo Aplicável do Outlook](https://docs.microsoft.com/outlook/add-ins/#extension-points)| Escrever ou Ler |

##### <a name="example"></a>Exemplo

```javascript
function myHandlerFunction(eventarg) {
  if (eventarg.attachmentStatus === Office.MailboxEnums.AttachmentStatus.Added) {
    var attachment = eventarg.attachmentDetails;
    console.log("Event Fired and Attachment Added!");
    getAttachmentContentAsync(attachment.id, options, callback);
  }
}

Office.context.mailbox.item.addHandlerAsync(Office.EventType.AttachmentsChanged, myHandlerFunction, myCallback);
```

####  <a name="additemattachmentasyncitemid-attachmentname-options-callback"></a>addItemAttachmentAsync(itemId, attachmentName, [options], [callback])

Adiciona um item do Exchange, como uma mensagem, como anexo na mensagem ou no compromisso.

O método `addItemAttachmentAsync` anexa o item com o identificador do Exchange especificado ao item no formulário de composição. Se você especificar um método de retorno de chamada, o método é chamado com um parâmetro, `asyncResult`, que contém o identificador do anexo ou um código que indica qualquer erro que tenha ocorrido ao anexar o item. Você pode usar o parâmetro `options` para passar informações de estado ao método de retorno de chamada, se necessário.

Posteriormente, você poderá usar o identificador com o método [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) para remover o anexo na mesma sessão.

Se o Suplemento do Office estiver em execução no Outlook Web App, o método `addItemAttachmentAsync` pode anexar itens a itens que não sejam aquele que você está editando; no entanto, isso não tem suporte e não é recomendado.

##### <a name="parameters"></a>Parâmetros

|Nome|Tipo|Atributos|Descrição|
|---|---|---|---|
|`itemId`|String||O identificador do Exchange do item a anexar. O comprimento máximo é de 100 caracteres.|
|`attachmentName`|Cadeia de caracteres||O assunto do item a ser anexado. O tamanho máximo é de 255 caracteres.|
|`options`|Object|&lt;opcional&gt;|Um objeto literal que contém uma ou mais das propriedades a seguir.|
|`options.asyncContext`|Objeto|&lt;opcional&gt;|Os desenvolvedores podem fornecer qualquer objeto que desejarem acessar no método de retorno de chamada.|
|`callback`|function|&lt;opcional&gt;|Quando o método for concluído, a função passada ao parâmetro `callback` é chamada com um único parâmetro, `asyncResult`, que é um objeto [`AsyncResult`](/javascript/api/office/office.asyncresult). <br/>Em caso de êxito, o identificador do anexo será fornecido na propriedade `asyncResult.value`.<br/>Se houver falha ao adicionar o anexo, o objeto `asyncResult` conterá um objeto `Error` que fornece uma descrição do erro.|

##### <a name="errors"></a>Erros

|Código de erro|Descrição|
|------------|-------------|
|`NumberOfAttachmentsExceeded`|A mensagem ou o compromisso tem muitos anexos.|

##### <a name="requirements"></a>Requisitos

|Requisito|Valor|
|---|---|
|[Versão do conjunto de requisitos mínimos da caixa de correio](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|1.1|
|[Nível de permissão mínimo](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|ReadWriteItem|
|[Modo do Outlook aplicável](https://docs.microsoft.com/outlook/add-ins/#extension-points)|Escrever|

##### <a name="example"></a>Exemplo

O exemplo a seguir adiciona um item existente do Outlook como um anexo com o nome `My Attachment`.

```javascript
function callback(result) {
  if (result.error) {
    console.log(result.error);
  } else {
    console.log("Attachment added");
  }
}

function addAttachment() {
  // EWS ID of item to attach (shortened for readability).
  var itemId = "AAMkADI1...AAA=";

  // The values in asyncContext can be accessed in the callback.
  var options = { 'asyncContext': { var1: 1, var2: 2 } };

  Office.context.mailbox.item.addItemAttachmentAsync(itemId, "My Attachment", options, callback);
}
```

####  <a name="close"></a>close()

Fecha o item atual que está sendo composto.

O comportamento do método `close` depende do estado atual do item que está sendo redigido. Se o item tiver alterações não salvas, o cliente solicitará que o usuário salve, descarte ou cancele a ação ao fechar.

> [!NOTE]
> No Outlook na Web, se o item é um compromisso e já foi salvo usando `saveAsync`, o usuário é solicitado a salvar, descartar ou cancelar mesmo se não tiver havido alterações desde que o item foi salvo pela última vez.

No cliente do Outlook para área de trabalho, se a mensagem for uma resposta embutida, o método `close` não terá efeito.

##### <a name="requirements"></a>Requisitos

|Requisito|Valor|
|---|---|
|[Versão do conjunto de requisitos mínimos da caixa de correio](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|1.3|
|[Nível de permissão mínimo](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|Restrito|
|[Modo do Outlook aplicável](https://docs.microsoft.com/outlook/add-ins/#extension-points)|Escrever|

#### <a name="displayreplyallformformdata-callback"></a>displayReplyAllForm(formData, [callback])

Exibe um formulário de resposta que inclui o remetente e todos os destinatários da mensagem selecionada ou o organizador e todos os participantes do compromisso selecionado.

> [!NOTE]
> Não há suporte para esse método no Outlook para iOS ou no Outlook para Android.

No Outlook Web App, o formulário de resposta é exibido como um formulário pop-out no modo de exibição de três colunas e um formulário pop-up no modo de exibição de uma ou duas colunas.

Se qualquer dos parâmetros da cadeia de caracteres exceder seu limite, `displayReplyAllForm` gera uma exceção.

Quando os anexos são especificados no parâmetro `formData.attachments`, o Outlook e o Outlook Web App tentam baixar todos os anexos e anexá-los ao formulário de resposta. Se a adição de anexos falhar, será exibido um erro na interface de usuário do formulário. Se isso não for possível, nenhuma mensagem de erro será apresentada.

##### <a name="parameters"></a>Parâmetros

|Nome|Tipo|Atributos|Descrição|
|---|---|---|---|
|`formData`|String &#124; Object||Uma cadeia de caracteres que contém texto e HTML e que representa o corpo do formulário de resposta. A cadeia de caracteres está limitada a 32 KB.<br/>**OU**<br/>Um objeto que contém os dados do corpo ou do anexo e uma função de retorno de chamada. O objeto é definido da maneira a seguir.|
|`formData.htmlBody`|String|&lt;opcional&gt;|Uma cadeia de caracteres que contém texto e HTML e que representa o corpo do formulário de resposta. A cadeia de caracteres está limitada a 32 KB.
|`formData.attachments`|Array.&lt;Object&gt;|&lt;opcional&gt;|Uma matriz de objetos JSON que são anexos de arquivo ou item.|
|`formData.attachments.type`|String||Indica o tipo de anexo. Deve ser `file` para um anexo de arquivo ou `item` para um anexo de item.|
|`formData.attachments.name`|String||Uma cadeia de caracteres que contém o nome do anexo, até 255 caracteres de comprimento.|
|`formData.attachments.url`|String||Usado somente se `type` estiver definido como `file`. O URI do local para o arquivo.|
|`formData.attachments.isInline`|Boolean||Usado somente se `type` estiver definido como `file`. Se for `true`, indicará que o anexo será mostrado embutido no corpo da mensagem e não deverá ser exibido na lista de anexos.|
|`formData.attachments.itemId`|Cadeia de caracteres||Usado somente se `type` estiver definido como `item`. A ID do item do EWS do anexo. Isso é uma cadeia de até 100 caracteres.|
|`callback`|function|&lt;opcional&gt;|Quando o método for concluído, a função passada ao parâmetro `callback` é chamada com um único parâmetro, `asyncResult`, que é um objeto [AsyncResult](/javascript/api/office/office.asyncresult).|

##### <a name="requirements"></a>Requisitos

|Requisito|Valor|
|---|---|
|[Versão do conjunto de requisitos mínimos da caixa de correio](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|1.0|
|[Nível de permissão mínimo](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|ReadItem|
|[Modo do Outlook aplicável](https://docs.microsoft.com/outlook/add-ins/#extension-points)|Read|

##### <a name="examples"></a>Exemplos

O código a seguir transmite uma cadeia de caracteres à função `displayReplyAllForm`.

```javascript
Office.context.mailbox.item.displayReplyAllForm('hello there');
Office.context.mailbox.item.displayReplyAllForm('<b>hello there</b>');
```

Responder com um corpo vazio.

```javascript
Office.context.mailbox.item.displayReplyAllForm({});
```

Responder apenas com um corpo.

```javascript
Office.context.mailbox.item.displayReplyAllForm(
{
  'htmlBody' : 'hi'
});
```

Responder com um corpo e um anexo de arquivo.

```javascript
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

```javascript
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

Responder com um corpo, um anexo de arquivo, um anexo do item e um retorno de chamada.

```javascript
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

#### <a name="displayreplyformformdata-callback"></a>displayReplyForm(formData, [callback])

Exibe um formulário de resposta que inclui o remetente da mensagem selecionada ou o organizador do compromisso selecionado.

> [!NOTE]
> Não há suporte para esse método no Outlook para iOS ou no Outlook para Android.

No Outlook Web App, o formulário de resposta é exibido como um formulário pop-out no modo de exibição de três colunas e um formulário pop-up no modo de exibição de uma ou duas colunas.

Se qualquer dos parâmetros da cadeia de caracteres exceder seu limite, `displayReplyForm` gera uma exceção.

Quando os anexos são especificados no parâmetro `formData.attachments`, o Outlook e o Outlook Web App tentam baixar todos os anexos e anexá-los ao formulário de resposta. Se a adição de anexos falhar, será exibido um erro na interface de usuário do formulário. Se isso não for possível, nenhuma mensagem de erro será apresentada.

##### <a name="parameters"></a>Parâmetros

|Nome|Tipo|Atributos|Descrição|
|---|---|---|---|
|`formData`|String &#124; Object||Uma cadeia de caracteres que contém texto e HTML e que representa o corpo do formulário de resposta. A cadeia de caracteres está limitada a 32 KB.<br/>**OU**<br/>Um objeto que contém os dados do corpo ou do anexo e uma função de retorno de chamada. O objeto é definido da maneira a seguir.|
|`formData.htmlBody`|String|&lt;opcional&gt;|Uma cadeia de caracteres que contém texto e HTML e que representa o corpo do formulário de resposta. A cadeia de caracteres está limitada a 32 KB.
|`formData.attachments`|Array.&lt;Object&gt;|&lt;opcional&gt;|Uma matriz de objetos JSON que são anexos de arquivo ou item.|
|`formData.attachments.type`|String||Indica o tipo de anexo. Deve ser `file` para um anexo de arquivo ou `item` para um anexo de item.|
|`formData.attachments.name`|String||Uma cadeia de caracteres que contém o nome do anexo, até 255 caracteres de comprimento.|
|`formData.attachments.url`|String||Usado somente se `type` estiver definido como `file`. O URI do local para o arquivo.|
|`formData.attachments.isInline`|Boolean||Usado somente se `type` estiver definido como `file`. Se for `true`, indicará que o anexo será mostrado embutido no corpo da mensagem e não deverá ser exibido na lista de anexos.|
|`formData.attachments.itemId`|Cadeia de caracteres||Usado somente se `type` estiver definido como `item`. A ID do item do EWS do anexo. Isso é uma cadeia de até 100 caracteres.|
|`callback`|function|&lt;opcional&gt;|Quando o método for concluído, a função passada ao parâmetro `callback` é chamada com um único parâmetro, `asyncResult`, que é um objeto [AsyncResult](/javascript/api/office/office.asyncresult).|

##### <a name="requirements"></a>Requisitos

|Requisito|Valor|
|---|---|
|[Versão do conjunto de requisitos mínimos da caixa de correio](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|1.0|
|[Nível de permissão mínimo](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|ReadItem|
|[Modo do Outlook aplicável](https://docs.microsoft.com/outlook/add-ins/#extension-points)|Read|

##### <a name="examples"></a>Exemplos

O código a seguir transmite uma cadeia de caracteres à função `displayReplyForm`.

```javascript
Office.context.mailbox.item.displayReplyForm('hello there');
Office.context.mailbox.item.displayReplyForm('<b>hello there</b>');
```

Responder com um corpo vazio.

```javascript
Office.context.mailbox.item.displayReplyForm({});
```

Responder apenas com um corpo.

```javascript
Office.context.mailbox.item.displayReplyForm(
{
  'htmlBody' : 'hi'
});
```

Responder com um corpo e um anexo de arquivo.

```javascript
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

```javascript
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

Responder com um corpo, um anexo de arquivo, um anexo do item e um retorno de chamada.

```javascript
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

#### <a name="getattachmentcontentasyncattachmentid-options-callback--attachmentcontentjavascriptapioutlookofficeattachmentcontent"></a>getAttachmentContentAsync (attachmentid, [opções], [callback]) → [AttachmentContent](/javascript/api/outlook/office.attachmentcontent)

Obtém o anexo especificado de uma mensagem ou compromisso e o retorna como um objeto `AttachmentContent`.

O método `getAttachmentContentAsync` remove o obtém anexo com o identificador especificado do item. Como melhor prática, você deve usar o identificador para recuperar um anexo na mesma sessão da qual attachmentIds foram recuperadas com o chamada `getAttachmentsAsync` ou `item.attachments`. No Outlook Web App e no OWA para Dispositivos, o identificador do anexo é válido apenas dentro da mesma sessão. Uma sessão é finalizada quando o usuário fecha o aplicativo, ou se o usuário começa a compor em um formulário embutido e, subsequentemente, sai do formulário para continuar em uma janela separada.

##### <a name="parameters"></a>Parâmetros

|Nome|Tipo|Atributos|Descrição|
|---|---|---|---|
|`attachmentId`|String||O identificador do anexo que você quer obter.|
|`options`|Objeto|&lt;opcional&gt;|Um objeto literal que contém uma ou mais das propriedades a seguir.|
|`options.asyncContext`|Objeto|&lt;opcional&gt;|Os desenvolvedores podem fornecer qualquer objeto que desejarem acessar no método de retorno de chamada.|
|`callback`|function|&lt;opcional&gt;|Quando o método for concluído, a função passada ao parâmetro `callback` é chamada com um único parâmetro, `asyncResult`, que é um objeto [AsyncResult](/javascript/api/office/office.asyncresult).|

##### <a name="requirements"></a>Requisitos

|Requisito|Valor|
|---|---|
|[Versão do conjunto de requisitos mínimos da caixa de correio](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|Visualização|
|[Nível de permissão mínimo](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|ReadItem|
|[Modo Aplicável do Outlook](https://docs.microsoft.com/outlook/add-ins/#extension-points)|Escrever ou Ler|

##### <a name="returns"></a>Retorna:

Tipo: [AttachmentContent](/javascript/api/outlook/office.attachmentcontent)

##### <a name="example"></a>Exemplo

```javascript
var item = Office.context.mailbox.item;
var listOfAttachments = [];
item.getAttachmentsAsync(callback);

function callback(result) {
  if (result.value.length > 0) {
    for (i = 0 ; i < result.value.length ; i++) {
      var options = {asyncContext: {type: result.value[i].attachmentType}};
      getAttachmentContentAsync(result.value[i].id, options, handleAttachmentsCallback);
    }
  }
}

function handleAttachmentsCallback(result) {
  // Parse string to be a url, an .eml file, a base64-encoded string, or an .icalendar file.
  if (result.format === Office.MailboxEnums.AttachmentContentFormat.Base64) {
    // Handle file attachment.
  } else if (result.format === Office.MailboxEnums.AttachmentContentFormat.Eml) {
    // Handle email item attachment.
  } else if (result.format === Office.MailboxEnums.AttachmentContentFormat.ICalendar) {
    // Handle .icalender attachment.
  } else if (result.format === Office.MailboxEnums.AttachmentContentFormat.Url) {
    // Handle cloud attachment.
  } else {
    // Handle attachment formats that are not supported.
  }
}
```

#### <a name="getattachmentsasyncoptions-callback--arrayattachmentdetailsjavascriptapioutlookofficeattachmentdetails"></a>getAttachmentsAsync ([opções], [callback]) → array. <[AttachmentDetails](/javascript/api/outlook/office.attachmentdetails)>

Obtém os anexos do item como uma matriz. Somente modo de redação.

##### <a name="parameters"></a>Parâmetros

|Nome|Tipo|Atributos|Descrição|
|---|---|---|---|
|`options`|Object|&lt;opcional&gt;|Um objeto literal que contém uma ou mais das propriedades a seguir.|
|`options.asyncContext`|Objeto|&lt;opcional&gt;|Os desenvolvedores podem fornecer qualquer objeto que desejarem acessar no método de retorno de chamada.|
|`callback`|function|&lt;opcional&gt;|Quando o método for concluído, a função passada ao parâmetro `callback` é chamada com um único parâmetro, `asyncResult`, que é um objeto [AsyncResult](/javascript/api/office/office.asyncresult).|

##### <a name="requirements"></a>Requisitos

|Requisito|Valor|
|---|---|
|[Versão do conjunto de requisitos mínimos da caixa de correio](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|Visualização|
|[Nível de permissão mínimo](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|ReadItem|
|[Modo Aplicável do Outlook](https://docs.microsoft.com/outlook/add-ins/#extension-points)|Escrever|

##### <a name="returns"></a>Retorna:

Tipo: Array.<[AttachmentDetails](/javascript/api/outlook/office.attachmentdetails)>

##### <a name="example"></a>Exemplo

O exemplo a seguir cria uma cadeia de caracteres HTML com detalhes de todos os anexos no item atual.

```javascript
var item = Office.context.mailbox.item;
var outputString = "";
item.getAttachmentsAsync(callback);

function callback(result) {
  if (result.value.length > 0) {
    for (i = 0 ; i < result.value.length ; i++) {
      var attachment = result.value [i];
      outputString += "<BR>" + i + ". Name: ";
      outputString += attachment.name;
      outputString += "<BR>ID: " + attachment.id;
      outputString += "<BR>contentType: " + attachment.contentType;
      outputString += "<BR>size: " + attachment.size;
      outputString += "<BR>attachmentType: " + attachment.attachmentType;
      outputString += "<BR>isInline: " + attachment.isInline;
    }
  }
}
```

#### <a name="getentities--entitiesjavascriptapioutlookofficeentities"></a>getEntities() → {[Entities](/javascript/api/outlook/office.entities)}

Obtém as entidades encontradas no corpo do item selecionado.

> [!NOTE]
> Não há suporte para esse método no Outlook para iOS ou no Outlook para Android.

##### <a name="requirements"></a>Requisitos

|Requisito|Valor|
|---|---|
|[Versão do conjunto de requisitos mínimos da caixa de correio](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|1.0|
|[Nível de permissão mínimo](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|ReadItem|
|[Modo Aplicável do Outlook](https://docs.microsoft.com/outlook/add-ins/#extension-points)|Read|

##### <a name="returns"></a>Retorna:

Tipo: [Entities](/javascript/api/outlook/office.entities)

##### <a name="example"></a>Exemplo

O exemplo a seguir acessa as entidades de contatos no corpo do item atual.

```javascript
var contacts = Office.context.mailbox.item.getEntities().contacts;
```

#### <a name="getentitiesbytypeentitytype--nullable-arraystringcontactjavascriptapioutlookofficecontactmeetingsuggestionjavascriptapioutlookofficemeetingsuggestionphonenumberjavascriptapioutlookofficephonenumbertasksuggestionjavascriptapioutlookofficetasksuggestion"></a>getEntitiesByType(entityType) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook/office.contact)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion))>}

Obtém uma matriz de todas as entidades do tipo de entidade especificado encontradas no corpo do item selecionado.

> [!NOTE]
> Não há suporte para esse método no Outlook para iOS ou no Outlook para Android.

##### <a name="parameters"></a>Parâmetros

|Nome|Tipo|Descrição|
|---|---|---|
|`entityType`|[Office.MailboxEnums.EntityType](/javascript/api/outlook/office.mailboxenums.entitytype)|Um dos valores de enumeração de EntityType.|

##### <a name="requirements"></a>Requisitos

|Requisito|Valor|
|---|---|
|[Versão do conjunto de requisitos mínimos da caixa de correio](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|1.0|
|[Nível de permissão mínimo](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|Restrito|
|[Modo do Outlook aplicável](https://docs.microsoft.com/outlook/add-ins/#extension-points)|Read|

##### <a name="returns"></a>Retorna:

Se o valor passado em `entityType` não for um membro válido da enumeração `EntityType`, o método retorna nulo. Se nenhuma entidade do tipo especificado estiver presente no corpo do item, o método retorna uma matriz vazia. Caso contrário, o tipo de objetos na matriz retornada depende do tipo de entidade solicitado no parâmetro `entityType`.

Enquanto o nível de permissão mínimo a usar esse método é **Restricted**, alguns tipos de entidade requerem **ReadItem** para obter acesso, conforme especificado na tabela a seguir.

|Valor de `entityType`|Tipo de objetos na matriz retornada|Nível de permissão exigido|
|---|---|---|
|`Address`|String|**Restrito**|
|`Contact`|Contato|**ReadItem**|
|`EmailAddress`|String|**ReadItem**|
|`MeetingSuggestion`|MeetingSuggestion|**ReadItem**|
|`PhoneNumber`|PhoneNumber|**Restrito**|
|`TaskSuggestion`|TaskSuggestion|**ReadItem**|
|`URL`|String|**Restrito**|

Tipo: Array.<(String|[Contact](/javascript/api/outlook/office.contact)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion))>

##### <a name="example"></a>Exemplo

O exemplo a seguir mostra como acessar uma matriz de cadeias de caracteres que representa endereços postais no corpo do item atual.

```javascript
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
};
```

#### <a name="getfilteredentitiesbynamename--nullable-arraystringcontactjavascriptapioutlookofficecontactmeetingsuggestionjavascriptapioutlookofficemeetingsuggestionphonenumberjavascriptapioutlookofficephonenumbertasksuggestionjavascriptapioutlookofficetasksuggestion"></a>getFilteredEntitiesByName(name) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook/office.contact)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion))>}

Retorna entidades bem conhecidas no item selecionado que passam o filtro nomeado definido no arquivo de manifesto XML.

> [!NOTE]
> Não há suporte para esse método no Outlook para iOS ou no Outlook para Android.

O método `getFilteredEntitiesByName` retorna as entidades que correspondem à expressão regular definida no elemento de regra [ItemHasKnownEntity](/office/dev/add-ins/reference/manifest/rule#itemhasknownentity-rule) no arquivo de manifesto XML com o valor do elemento `FilterName` especificado.

##### <a name="parameters"></a>Parâmetros

|Nome|Tipo|Descrição|
|---|---|---|
|`name`|String|O nome do elemento de regra `ItemHasKnownEntity` que define o filtro a corresponder.|

##### <a name="requirements"></a>Requisitos

|Requisito|Valor|
|---|---|
|[Versão do conjunto de requisitos mínimos da caixa de correio](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|1.0|
|[Nível de permissão mínimo](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|ReadItem|
|[Modo do Outlook aplicável](https://docs.microsoft.com/outlook/add-ins/#extension-points)|Read|

##### <a name="returns"></a>Retorna:

Se não houver nenhum elemento `ItemHasKnownEntity` no manifesto com um valor de elemento `FilterName` que corresponda ao parâmetro `name`, o método retorna `null`. Se o parâmetro `name` corresponder a um elemento `ItemHasKnownEntity` no manifesto, mas não houver entidades no item atual que correspondam, o método retorna uma matriz vazia.

Tipo: Array.<(String|[Contact](/javascript/api/outlook/office.contact)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion))>

#### <a name="getinitializationcontextasyncoptions-callback"></a>getInitializationContextAsync([options], [callback])

Obtém dados de inicialização que são transmitidos quando o suplemento é [ativado por uma mensagem acionável](https://docs.microsoft.com/outlook/actionable-messages/invoke-add-in-from-actionable-message).

> [!NOTE]
> Esse método só é compatível com o Outlook 2016 ou posterior para Windows (versões Clique para Executar posteriores à 16.0.8413.1000) e o Outlook na Web para o Office 365.

##### <a name="parameters"></a>Parâmetros
|Nome|Tipo|Atributos|Descrição|
|---|---|---|---|
|`options`|Objeto|&lt;opcional&gt;|Um objeto literal que contém uma ou mais das propriedades a seguir.|
|`options.asyncContext`|Objeto|&lt;opcional&gt;|Os desenvolvedores podem fornecer qualquer objeto que desejarem acessar no método de retorno de chamada.|
|`callback`|function|&lt;opcional&gt;|Quando o método for concluído, a função passada ao parâmetro `callback` é chamada com um único parâmetro, `asyncResult`, que é um objeto [`AsyncResult`](/javascript/api/office/office.asyncresult). <br/>Após o êxito, os dados de inicialização são fornecidos na propriedade `asyncResult.value` como uma cadeia de caracteres.<br/>Se não houver nenhum contexto de inicialização, o objeto `asyncResult` conterá um objeto `Error` com sua propriedade `code` definida como `9020` e sua propriedade `name` definida como `GenericResponseError`.|

##### <a name="requirements"></a>Requisitos

|Requisito|Valor|
|---|---|
|[Versão do conjunto de requisitos mínimos da caixa de correio](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|Visualização|
|[Nível de permissão mínimo](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|ReadItem|
|[Modo do Outlook aplicável](https://docs.microsoft.com/outlook/add-ins/#extension-points)|Leitura|

##### <a name="example"></a>Exemplo

```javascript
// Get the initialization context (if present).
Office.context.mailbox.item.getInitializationContextAsync(
  function(asyncResult) {
    if (asyncResult.status == Office.AsyncResultStatus.Succeeded) {
      if (asyncResult.value != null && asyncResult.value.length > 0) {
        // The value is a string, parse to an object.
        var context = JSON.parse(asyncResult.value);
        // Do something with context.
      } else {
        // Empty context, treat as no context.
      }
    } else {
      if (asyncResult.error.code == 9020) {
        // GenericResponseError returned when there is no context.
        // Treat as no context.
      } else {
        // Handle the error.
      }
    }
  }
);
```

#### <a name="getregexmatches--object"></a>getRegExMatches() → {Object}

Retorna valores de cadeia de caracteres ao item selecionado que correspondem às expressões regulares definidas no arquivo de manifesto XML.

> [!NOTE]
> Não há suporte para esse método no Outlook para iOS ou no Outlook para Android.

O método `getRegExMatches` retorna as cadeias de caracteres que correspondem à expressão regular definida em cada elemento de regra `ItemHasRegularExpressionMatch` ou `ItemHasKnownEntity` no arquivo de manifesto XML. Para uma regra `ItemHasRegularExpressionMatch`, uma cadeia de caracteres correspondente deve ocorrer na propriedade do item especificada por essa regra. O tipo simples `PropertyName` define as propriedades compatíveis.

Por exemplo, considere que um manifesto de suplemento tem o seguinte elemento `Rule`:

```xml
<Rule xsi:type="RuleCollection" Mode="And">
  <Rule xsi:type="ItemIs" FormType="Read" ItemType="Message" />
  <Rule xsi:type="RuleCollection" Mode="Or">
    <Rule xsi:type="ItemHasRegularExpressionMatch" RegExName="fruits" RegExValue="apple|banana|coconut" PropertyName="BodyAsPlaintext" IgnoreCase="true" />
    <Rule xsi:type="ItemHasRegularExpressionMatch" RegExName="veggies" RegExValue="tomato|onion|spinach|broccoli" PropertyName="BodyAsPlaintext" IgnoreCase="true" />
  </Rule>
</Rule>
```

O objeto retornado por `getRegExMatches` teria duas propriedades: `fruits` e `veggies`.

```json
{
  'fruits': ['apple','banana','Banana','coconut'],
  'veggies': ['tomato','onion','spinach','broccoli']
}
```

Se você especificar uma regra `ItemHasRegularExpressionMatch` na propriedade do corpo de um item, a expressão regular deverá filtrar mais o corpo e não tentar retornar todo o corpo do item. Usar uma expressão regular como `.*` para obter todo o corpo de um item nem sempre retorna os resultados esperados. Em vez disso, use o método [`Body.getAsync`](/javascript/api/outlook/office.body#getasync-coerciontype--options--callback-) para recuperar todo o corpo.

##### <a name="requirements"></a>Requisitos

|Requisito|Valor|
|---|---|
|[Versão do conjunto de requisitos mínimos da caixa de correio](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|1.0|
|[Nível de permissão mínimo](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|ReadItem|
|[Modo do Outlook aplicável](https://docs.microsoft.com/outlook/add-ins/#extension-points)|Read|

##### <a name="returns"></a>Retorna:

Um objeto que contém matrizes de cadeias de caracteres que correspondem às expressões regulares definidas no arquivo XML do manifesto. O nome de cada matriz é igual ao valor correspondente do atributo `RegExName` da regra `ItemHasRegularExpressionMatch` correspondente ou do atributo `FilterName` da regra `ItemHasKnownEntity` correspondente.

<dl class="param-type">

<dt>Type</dt>

<dd>Objeto</dd>

</dl>

##### <a name="example"></a>Exemplo

O exemplo a seguir mostra como acessar a matriz de correspondências para os elementos de regra de expressão regular `fruits` e `veggies`, que estão especificados no manifesto.

```javascript
var allMatches = Office.context.mailbox.item.getRegExMatches();
var fruits = allMatches.fruits;
var veggies = allMatches.veggies;
```

#### <a name="getregexmatchesbynamename--nullable-array-string-"></a>getRegExMatchesByName(name) → (nullable) {Array.< String >}

Retorna valores de cadeia de caracteres no item selecionado que correspondem à expressão regular nomeada definida no arquivo de manifesto XML.

> [!NOTE]
> Não há suporte para esse método no Outlook para iOS ou no Outlook para Android.

O método `getRegExMatchesByName` retorna as cadeias de caracteres que correspondem à expressão regular definida no elemento de regra `ItemHasRegularExpressionMatch` no arquivo de manifesto XML com o valor de elemento `RegExName` especificado.

Se você especificar uma regra `ItemHasRegularExpressionMatch` na propriedade do corpo de um item, a expressão regular deverá filtrar mais o corpo e não tentar retornar todo o corpo do item. Usar uma expressão regular como `.*` para obter todo o corpo de um item nem sempre retorna os resultados esperados.

##### <a name="parameters"></a>Parâmetros

|Nome|Tipo|Descrição|
|---|---|---|
|`name`|String|O nome do elemento de regra `ItemHasRegularExpressionMatch` que define o filtro a corresponder.|

##### <a name="requirements"></a>Requisitos

|Requisito|Valor|
|---|---|
|[Versão do conjunto de requisitos mínimos da caixa de correio](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|1.0|
|[Nível de permissão mínimo](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|ReadItem|
|[Modo do Outlook aplicável](https://docs.microsoft.com/outlook/add-ins/#extension-points)|Read|

##### <a name="returns"></a>Retorna:

Uma matriz que contém as cadeias de caracteres que correspondem à expressão regular definida no arquivo de manifesto XML.

<dl class="param-type">

<dt>Type</dt>

<dd>Array.< String ></dd>

</dl>

##### <a name="example"></a>Exemplo

```javascript
var fruits = Office.context.mailbox.item.getRegExMatchesByName("fruits");
var veggies = Office.context.mailbox.item.getRegExMatchesByName("veggies");
```

####  <a name="getselecteddataasynccoerciontype-options-callback--string"></a>getSelectedDataAsync(coercionType, [options], callback) → {String}

Retorna de forma assíncrona os dados selecionados do assunto ou do corpo de uma mensagem.

Se não houver seleção, mas o cursor estiver no corpo ou no assunto, o método retorna nulo para os dados selecionados. Se um campo que não seja o corpo ou o assunto estiver selecionado, o método retorna o erro `InvalidSelection`.

##### <a name="parameters"></a>Parâmetros

|Nome|Tipo|Atributos|Descrição|
|---|---|---|---|
|`coercionType`|[Office.CoercionType](office.md#coerciontype-string)||Solicita um formato para os dados. Se Text, o método retorna o texto sem formatação como uma cadeia de caracteres, removendo quaisquer marcas HTML presentes. Se HTML, o método retorna o texto selecionado, seja ele texto sem formatação ou HTML.|
|`options`|Objeto|&lt;opcional&gt;|Um objeto literal que contém uma ou mais das propriedades a seguir.|
|`options.asyncContext`|Objeto|&lt;opcional&gt;|Os desenvolvedores podem fornecer qualquer objeto que desejarem acessar no método de retorno de chamada.|
|`callback`|function||Quando o método for concluído, a função passada ao parâmetro `callback` será chamada com um único parâmetro, `asyncResult`, que é um objeto [`AsyncResult`](/javascript/api/office/office.asyncresult).<br/><br/>Para acessar os dados selecionados do método de retorno de chamada, chame `asyncResult.value.data`. Para acessar a propriedade de origem de que a seleção é proveniente, chame `asyncResult.value.sourceProperty`, que será `body` ou `subject`.|

##### <a name="requirements"></a>Requisitos

|Requisito|Valor|
|---|---|
|[Versão do conjunto de requisitos mínimos da caixa de correio](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|1.2|
|[Nível de permissão mínimo](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|ReadWriteItem|
|[Modo do Outlook aplicável](https://docs.microsoft.com/outlook/add-ins/#extension-points)|Escrever|

##### <a name="returns"></a>Retorna:

Os dados selecionados como uma cadeia de caracteres com formato determinado por `coercionType`.

<dl class="param-type">

<dt>Type</dt>

<dd>String</dd>

</dl>

##### <a name="example"></a>Exemplo

```javascript
// Get selected data.
Office.initialize = function () {
  Office.context.mailbox.item.getSelectedDataAsync(Office.CoercionType.Text, {}, getCallback);
};

function getCallback(asyncResult) {
  var text = asyncResult.value.data;
  var prop = asyncResult.value.sourceProperty;

  Office.context.mailbox.item.setSelectedDataAsync('Setting ' + prop + ': ' + text, {}, setCallback);
}

function setCallback(asyncResult) {
  // Check for errors.
}
```

#### <a name="getselectedentities--entitiesjavascriptapioutlookofficeentities"></a>getSelectedEntities() → {[Entities](/javascript/api/outlook/office.entities)}

Obtém as entidades encontradas em uma correspondência realçada que um usuário selecionou. As correspondências realçadas aplicam-se a [suplementos contextuais](https://docs.microsoft.com/outlook/add-ins/contextual-outlook-add-ins).

> [!NOTE]
> Não há suporte para esse método no Outlook para iOS ou no Outlook para Android.

##### <a name="requirements"></a>Requisitos

|Requisito|Valor|
|---|---|
|[Versão do conjunto de requisitos mínimos da caixa de correio](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|1.6|
|[Nível de permissão mínimo](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|ReadItem|
|[Modo do Outlook aplicável](https://docs.microsoft.com/outlook/add-ins/#extension-points)|Read|

##### <a name="returns"></a>Retorna:

Tipo: [Entities](/javascript/api/outlook/office.entities)

##### <a name="example"></a>Exemplo

O exemplo a seguir acessa as entidades de endereços na correspondência realçada, selecionada pelo usuário.

```javascript
var contacts = Office.context.mailbox.item.getSelectedEntities().addresses;
```

#### <a name="getselectedregexmatches--object"></a>getSelectedRegExMatches() → {Object}

Retorna valores de cadeia de caracteres em uma correspondência realçada que corresponde às expressões regulares definidas no arquivo de manifesto XML. As correspondências realçadas aplicam-se a [suplementos contextuais](https://docs.microsoft.com/outlook/add-ins/contextual-outlook-add-ins).

> [!NOTE]
> Não há suporte para esse método no Outlook para iOS ou no Outlook para Android.

O método `getSelectedRegExMatches` retorna as cadeias de caracteres que correspondem à expressão regular definida em cada elemento de regra `ItemHasRegularExpressionMatch` ou `ItemHasKnownEntity` no arquivo de manifesto XML. Para uma regra `ItemHasRegularExpressionMatch`, uma cadeia de caracteres correspondente deve ocorrer na propriedade do item especificada por essa regra. O tipo simples `PropertyName` define as propriedades compatíveis.

Por exemplo, considere que um manifesto de suplemento tem o seguinte elemento `Rule`:

```xml
<Rule xsi:type="RuleCollection" Mode="And">
  <Rule xsi:type="ItemIs" FormType="Read" ItemType="Message" />
  <Rule xsi:type="RuleCollection" Mode="Or">
    <Rule xsi:type="ItemHasRegularExpressionMatch" RegExName="fruits" RegExValue="apple|banana|coconut" PropertyName="BodyAsPlaintext" IgnoreCase="true" />
    <Rule xsi:type="ItemHasRegularExpressionMatch" RegExName="veggies" RegExValue="tomato|onion|spinach|broccoli" PropertyName="BodyAsPlaintext" IgnoreCase="true" />
  </Rule>
</Rule>
```

O objeto retornado por `getRegExMatches` teria duas propriedades: `fruits` e `veggies`.

```json
{
  'fruits': ['apple','banana','Banana','coconut'],
  'veggies': ['tomato','onion','spinach','broccoli']
}
```

Se você especificar uma regra `ItemHasRegularExpressionMatch` na propriedade do corpo de um item, a expressão regular deverá filtrar mais o corpo e não tentar retornar todo o corpo do item. Usar uma expressão regular como `.*` para obter todo o corpo de um item nem sempre retorna os resultados esperados. Em vez disso, use o método [`Body.getAsync`](/javascript/api/outlook/office.body#getasync-coerciontype--options--callback-) para recuperar todo o corpo.

##### <a name="requirements"></a>Requisitos

|Requisito|Valor|
|---|---|
|[Versão do conjunto de requisitos mínimos da caixa de correio](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|1.6|
|[Nível de permissão mínimo](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|ReadItem|
|[Modo do Outlook aplicável](https://docs.microsoft.com/outlook/add-ins/#extension-points)|Read|

##### <a name="returns"></a>Retorna:

Um objeto que contém matrizes de cadeias de caracteres que correspondem às expressões regulares definidas no arquivo XML do manifesto. O nome de cada matriz é igual ao valor correspondente do atributo `RegExName` da regra `ItemHasRegularExpressionMatch` correspondente ou do atributo `FilterName` da regra `ItemHasKnownEntity` correspondente.

##### <a name="example"></a>Exemplo

O exemplo a seguir mostra como acessar a matriz de correspondências para os elementos de regra de expressão regular `fruits` e `veggies`, que estão especificados no manifesto.

```javascript
var selectedMatches = Office.context.mailbox.item.getSelectedRegExMatches();
var fruits = selectedMatches.fruits;
var veggies = selectedMatches.veggies;
```

#### <a name="getsharedpropertiesasyncoptions-callback"></a>getSharedPropertiesAsync([options], callback)

Obtém as propriedades do compromisso ou mensagem selecionada em uma pasta compartilhada, calendário ou caixa de correio.

##### <a name="parameters"></a>Parâmetros

|Nome|Tipo|Atributos|Descrição|
|---|---|---|---|
|`options`|Object|&lt;opcional&gt;|Um objeto literal que contém uma ou mais das propriedades a seguir.|
|`options.asyncContext`|Objeto|&lt;opcional&gt;|Os desenvolvedores podem fornecer qualquer objeto que desejarem acessar no método de retorno de chamada.|
|`callback`|function||Quando o método for concluído, a função passada ao parâmetro `callback` é chamada com um único parâmetro, `asyncResult`, que é um objeto [`AsyncResult`](/javascript/api/office/office.asyncresult).<br/><br/>As propriedades compartilhadas são fornecidas como um objeto [`SharedProperties`](/javascript/api/outlook/office.sharedproperties) na propriedade `asyncResult.value`. Esse objeto pode ser usado para obter as propriedades compartilhadas do item.|

##### <a name="requirements"></a>Requisitos

|Requisito|Valor|
|---|---|
|[Versão do conjunto de requisitos mínimos da caixa de correio](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|Visualização|
|[Nível de permissão mínimo](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|ReadItem|
|[Modo do Outlook aplicável](https://docs.microsoft.com/outlook/add-ins/#extension-points)|Escrever ou Ler|

##### <a name="example"></a>Exemplo

```javascript
Office.context.mailbox.item.getSharedPropertiesAsync(callback);

function callback (asyncResult) {
  var context = asyncResult.context;
  var sharedProperties = asyncResult.value;
}
```

####  <a name="loadcustompropertiesasynccallback-usercontext"></a>loadCustomPropertiesAsync(callback, [userContext])

Carrega de forma assíncrona as propriedades personalizadas para esse suplemento no item selecionado.

Propriedades personalizadas são armazenadas como pares chave/valor de acordo com o aplicativo e o item. Este método retorna um objeto `CustomProperties` no retorno de chamada, que oferece métodos para acessar as propriedades personalizadas específicas para o item atual e o suplemento atual. Propriedades personalizadas não são criptografadas no item, portanto não devem ser usadas como armazenamento seguro.

##### <a name="parameters"></a>Parâmetros

|Nome|Tipo|Atributos|Descrição|
|---|---|---|---|
|`callback`|function||Quando o método for concluído, a função passada ao parâmetro `callback` é chamada com um único parâmetro, `asyncResult`, que é um objeto [`AsyncResult`](/javascript/api/office/office.asyncresult).<br/><br/>As propriedades personalizadas são fornecidas como um objeto [`CustomProperties`](/javascript/api/outlook/office.customproperties) na propriedade `asyncResult.value`. Esse objeto pode ser usado para obter, definir e remover as propriedades personalizadas do item e salvar as alterações na propriedade personalizada definida de volta no servidor.|
|`userContext`|Objeto|&lt;opcional&gt;|Os desenvolvedores podem fornecer qualquer objeto que desejarem acessar na função de retorno de chamada. Esse objeto pode ser acessado pela propriedade `asyncResult.asyncContext` na função de retorno de chamada.|

##### <a name="requirements"></a>Requisitos

|Requisito|Valor|
|---|---|
|[Versão do conjunto de requisitos mínimos da caixa de correio](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|1.0|
|[Nível de permissão mínimo](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|ReadItem|
|[Modo do Outlook aplicável](https://docs.microsoft.com/outlook/add-ins/#extension-points)|Escrever ou Ler|

##### <a name="example"></a>Exemplo

O exemplo de código a seguir mostra como usar o método `loadCustomPropertiesAsync` para carregar de forma assíncrona as propriedades personalizadas que são específicas para o item atual. O exemplo também mostra como usar o método `CustomProperties.saveAsync` para salvar essas propriedades de volta no servidor. Depois de carregar as propriedades personalizadas, o exemplo de código usará o método `CustomProperties.get` para ler a propriedade personalizada `myProp`, o método `CustomProperties.set` para gravar na propriedade personalizada `otherProp` e, então, chama o método `saveAsync` para salvar as propriedades personalizadas.

```javascript
// The initialize function is required for all add-ins.
Office.initialize = function () {
  // Checks for the DOM to load using the jQuery ready function.
  $(document).ready(function () {
    // After the DOM is loaded, add-in-specific code can run.
    var item = Office.context.mailbox.item;
    item.loadCustomPropertiesAsync(customPropsCallback);
  });
};

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

O método `removeAttachmentAsync` remove o anexo com o identificador especificado do item. Como prática recomendada, deve-se usar o identificador do anexo para remover um anexo somente se o mesmo aplicativo de email tiver adicionado esse anexo na mesma sessão. No Outlook Web App e no OWA para Dispositivos, o identificador do anexo é válido apenas dentro da mesma sessão. Uma sessão é finalizada quando o usuário fecha o aplicativo, ou se o usuário começa a compor em um formulário embutido e, subsequentemente, sai do formulário para continuar em uma janela separada.

##### <a name="parameters"></a>Parâmetros

|Nome|Tipo|Atributos|Descrição|
|---|---|---|---|
|`attachmentId`|String||O identificador do anexo a remover.|
|`options`|Objeto|&lt;opcional&gt;|Um objeto literal que contém uma ou mais das propriedades a seguir.|
|`options.asyncContext`|Objeto|&lt;opcional&gt;|Os desenvolvedores podem fornecer qualquer objeto que desejarem acessar no método de retorno de chamada.|
|`callback`|function|&lt;opcional&gt;|Quando o método for concluído, a função transmitida ao parâmetro `callback` é chamada com um único parâmetro, `asyncResult`, que é um objeto [`AsyncResult`](/javascript/api/office/office.asyncresult). <br/>Se a remoção do anexo falhar, a propriedade `asyncResult.error` conterá um código de erro com o motivo da falha.|

##### <a name="errors"></a>Erros

|Código de erro|Descrição|
|------------|-------------|
|`InvalidAttachmentId`|O identificador de anexo não existe.|

##### <a name="requirements"></a>Requisitos

|Requisito|Valor|
|---|---|
|[Versão do conjunto de requisitos mínimos da caixa de correio](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|1.1|
|[Nível de permissão mínimo](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|ReadWriteItem|
|[Modo do Outlook aplicável](https://docs.microsoft.com/outlook/add-ins/#extension-points)|Escrever|

##### <a name="example"></a>Exemplo

O código a seguir remove um anexo com um identificador '0'.

```javascript
Office.context.mailbox.item.removeAttachmentAsync(
  '0',
  { asyncContext : null },
  function (asyncResult)
  {
    console.log(asyncResult.status);
  }
);
```

####  <a name="removehandlerasynceventtype-options-callback"></a>removeHandlerAsync(eventType, handler, [options], [callback])

Remove um manipulador de eventos para um tipo de evento com suporte.

Atualmente, os tipos de eventos `Office.EventType.AttachmentsChanged`suportados `Office.EventType.AppointmentTimeChanged`são `Office.EventType.EnhancedLocationsChanged`, `Office.EventType.RecipientsChanged`,, `Office.EventType.RecurrenceChanged`e.

##### <a name="parameters"></a>Parâmetros

| Nome | Tipo | Atributos | Descrição |
|---|---|---|---|
| `eventType` | [Office.EventType](office.md#eventtype-string) || O evento que deve revogar o manipulador. |
| `options` | Objeto | &lt;opcional&gt; | Um objeto literal que contém uma ou mais das propriedades a seguir. |
| `options.asyncContext` | Objeto | &lt;opcional&gt; | Os desenvolvedores podem fornecer qualquer objeto que desejarem acessar no método de retorno de chamada. |
| `callback` | function| &lt;opcional&gt;|Quando o método for concluído, a função transmitida ao parâmetro `callback` é chamada com um único parâmetro, `asyncResult`, que é um objeto [`AsyncResult`](/javascript/api/office/office.asyncresult).|

##### <a name="requirements"></a>Requisitos

|Requisito| Valor|
|---|---|
|[Versão do conjunto de requisitos mínimos da caixa de correio](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| 1.7 |
|[Nível de permissão mínimo](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| ReadItem |
|[Modo do Outlook aplicável](https://docs.microsoft.com/outlook/add-ins/#extension-points)| Escrever ou Ler |

####  <a name="saveasyncoptions-callback"></a>saveAsync([options], callback)

Salva um item de forma assíncrona.

Quando chamado, este método salva a mensagem atual como um rascunho e retorna a identificação do item por meio do método de retorno de chamada. No Outlook Web App ou no Outlook no modo online, o item é salvo no servidor. No Outlook no modo cache, o item é salvo no cache local.

> [!NOTE]
> Se seu suplemento chamar o `saveAsync` em um item no modo de redação a fim de obter um `itemId` para usar com a API EWS ou REST, esteja ciente de que quando o Outlook está no modo de cache, pode levar alguns instantes até que o item esteja realmente sincronizado com o servidor. Até que o item esteja sincronizado, usar o `itemId` retornará um erro.

Como compromissos não têm um estado de rascunho, se `saveAsync` for chamado em um compromisso no modo Redigir, o item será salvo como um compromisso normal no calendário do usuário. Para novos compromissos que não foram salvos antes, nenhum convite será enviado. Salvar um compromisso existente enviará uma atualização aos participantes adicionados ou removidos.

> [!NOTE]
> Os clientes a seguir têm diferentes comportamentos para `saveAsync` nos compromissos no modo de redação:
>
> - O Outlook para Mac não dá suporte ao `saveAsync` em uma reunião no modo composto. Chamar `saveAsync` em uma reunião no Outlook para Mac fará com que um erro seja retornado.
> - O Outlook na Web sempre envia um convite ou atualização quando `saveAsync` é chamado em um compromisso no modo de redação.

##### <a name="parameters"></a>Parâmetros

|Nome|Tipo|Atributos|Descrição|
|---|---|---|---|
|`options`|Object|&lt;opcional&gt;|Um objeto literal que contém uma ou mais das propriedades a seguir.|
|`options.asyncContext`|Objeto|&lt;opcional&gt;|Os desenvolvedores podem fornecer qualquer objeto que desejarem acessar no método de retorno de chamada.|
|`callback`|function||Quando o método for concluído, a função passada ao parâmetro `callback` é chamada com um único parâmetro, `asyncResult`, que é um objeto [`AsyncResult`](/javascript/api/office/office.asyncresult).<br/><br/>Em caso de sucesso, o identificador do item é fornecido na propriedade `asyncResult.value`.|

##### <a name="requirements"></a>Requisitos

|Requisito|Valor|
|---|---|
|[Versão do conjunto de requisitos mínimos da caixa de correio](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|1.3|
|[Nível de permissão mínimo](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|ReadWriteItem|
|[Modo do Outlook aplicável](https://docs.microsoft.com/outlook/add-ins/#extension-points)|Escrever|

##### <a name="examples"></a>Exemplos

```javascript
Office.context.mailbox.item.saveAsync(
  function callback(result) {
    // Process the result.
  });
```

A seguir apresentamos um exemplo do parâmetro `result` passado à função de retorno de chamada. A propriedade `value` contém a ID para o item.

```json
{
  "value":"AAMkADI5...AAA=",
  "status":"succeeded"
}
```

####  <a name="setselecteddataasyncdata-options-callback"></a>setSelectedDataAsync(data, [options], callback)

Insere de forma assíncrona os dados no corpo ou no assunto de uma mensagem.

O método `setSelectedDataAsync` insere a cadeia de caracteres especificada no local do cursor no corpo ou assunto do item ou, se o texto estiver selecionado no editor, substitui o texto selecionado. Se o cursor não estiver no campo do corpo ou assunto, um erro será retornado. Após a inserção, o cursor é colocado no final do conteúdo inserido.

##### <a name="parameters"></a>Parâmetros

|Nome|Tipo|Atributos|Descrição|
|---|---|---|---|
|`data`|String||Os dados a serem inseridos. Os dados não devem exceder 1.000.000 de caracteres. Se forem passados mais de 1.000.000 de caracteres, ocorrerá uma exceção `ArgumentOutOfRange`.|
|`options`|Objeto|&lt;opcional&gt;|Um objeto literal que contém uma ou mais das propriedades a seguir.|
|`options.asyncContext`|Objeto|&lt;opcional&gt;|Os desenvolvedores podem fornecer qualquer objeto que desejarem acessar no método de retorno de chamada.|
|`options.coercionType`|[Office.CoercionType](office.md#coerciontype-string)|&lt;opcional&gt;|Se `text`, o estilo atual é aplicado no Outlook Web App e no Outlook. Se o campo for um editor de HTML, apenas os dados de texto são inseridos, mesmo se os dados forem HTML.<br/><br/>Se `html` e o campo forem compatíveis com HTML (e o assunto não), o estilo atual é aplicado no Outlook Web App e o estilo padrão será aplicado no Outlook. Se o campo for um campo de texto, retorna um erro `InvalidDataFormat`.<br/><br/>Se `coercionType` não estiver definido, o resultado depende do campo: se o campo for HTML, HTML será usado; se o campo for texto, texto sem formatação será usado.|
|`callback`|function||Quando o método for concluído, a função passada ao parâmetro `callback` é chamada com um único parâmetro, `asyncResult`, que é um objeto [`AsyncResult`](/javascript/api/office/office.asyncresult).|

##### <a name="requirements"></a>Requisitos

|Requisito|Valor|
|---|---|
|[Versão do conjunto de requisitos mínimos da caixa de correio](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|1.2|
|[Nível de permissão mínimo](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|ReadWriteItem|
|[Modo do Outlook aplicável](https://docs.microsoft.com/outlook/add-ins/#extension-points)|Escrever|

##### <a name="example"></a>Exemplo

```javascript
Office.context.mailbox.item.setSelectedDataAsync("Hello World!");
Office.context.mailbox.item.setSelectedDataAsync("<b>Hello World!</b>", { coercionType : "html" });
```
