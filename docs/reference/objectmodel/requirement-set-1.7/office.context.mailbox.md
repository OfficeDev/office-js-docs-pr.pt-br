---
title: Office. Context. Mailbox – conjunto de requisitos 1,7
description: ''
ms.date: 06/20/2019
localization_priority: Normal
ms.openlocfilehash: 2cbe219e051eff63e27995e07a288ee787eeaf99
ms.sourcegitcommit: 3f5d7f4794e3d3c8bc3a79fa05c54157613b9376
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 08/02/2019
ms.locfileid: "36064512"
---
# <a name="mailbox"></a>mailbox

### <a name="officeofficemdcontextofficecontextmdmailbox"></a>[Office](Office.md)[.context](Office.context.md).mailbox

Fornece acesso ao modelo de objeto do suplemento do Outlook para o Microsoft Outlook.

##### <a name="requirements"></a>Requisitos

|Requisito| Valor|
|---|---|
|[Versão do conjunto de requisitos mínimos da caixa de correio](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| 1.0|
|[Nível de permissão mínimo](/outlook/add-ins/understanding-outlook-add-in-permissions)| Restrito|
|[Modo do Outlook aplicável](/outlook/add-ins/#extension-points)| Escrever ou Ler|

##### <a name="members-and-methods"></a>Membros e métodos

| Membro | Tipo |
|--------|------|
| [ewsUrl](#ewsurl-string) | Membro |
| [restUrl](#resturl-string) | Membro |
| [addHandlerAsync](#addhandlerasynceventtype-handler-options-callback) | Método |
| [convertToEwsId](#converttoewsiditemid-restversion--string) | Método |
| [convertToLocalClientTime](#converttolocalclienttimetimevalue--localclienttime) | Método |
| [convertToRestId](#converttorestiditemid-restversion--string) | Método |
| [convertToUtcClientTime](#converttoutcclienttimeinput--date) | Método |
| [displayAppointmentForm](#displayappointmentformitemid) | Método |
| [displayMessageForm](#displaymessageformitemid) | Método |
| [displayNewAppointmentForm](#displaynewappointmentformparameters) | Método |
| [displayNewMessageForm](#displaynewmessageformparameters) | Método |
| [getCallbackTokenAsync](#getcallbacktokenasyncoptions-callback) | Método |
| [getCallbackTokenAsync](#getcallbacktokenasynccallback-usercontext) | Método |
| [getUserIdentityTokenAsync](#getuseridentitytokenasynccallback-usercontext) | Método |
| [makeEwsRequestAsync](#makeewsrequestasyncdata-callback-usercontext) | Método |
| [removeHandlerAsync](#removehandlerasynceventtype-options-callback) | Método |

### <a name="namespaces"></a>Namespaces

[diagnostics](Office.context.mailbox.diagnostics.md): Fornece informações de diagnóstico para um suplemento do Outlook.

[item](Office.context.mailbox.item.md): Fornece propriedades e métodos para acessar uma mensagem ou um compromisso em um suplemento do Outlook.

[userProfile](Office.context.mailbox.userProfile.md): Fornece informações sobre o usuário em um suplemento do Outlook.

### <a name="members"></a>Membros

#### <a name="ewsurl-string"></a>ewsUrl: cadeia de caracteres

Obtém a URL do ponto de extremidade dos Serviços Web do Exchange (EWS) para esta conta de email. Somente modo de leitura.

> [!NOTE]
> Não há suporte para esse membro no Outlook no iOS ou no Android.

O valor `ewsUrl` pode ser usado por um serviço remoto para fazer chamadas do EWS à caixa de correio do usuário. Por exemplo, você pode criar um serviço remoto para [obter anexos do item selecionado](/outlook/add-ins/get-attachments-of-an-outlook-item).

Seu aplicativo deve ter a permissão **ReadItem** especificada em seu manifesto para chamar o membro `ewsUrl` em modo de leitura.

No modo de composição, é preciso chamar o método [`saveAsync`](Office.context.mailbox.item.md#saveasyncoptions-callback) antes de poder usar o membro `ewsUrl`. Seu aplicativo deve ter permissões **ReadWriteItem** para chamar o método `saveAsync`.

##### <a name="type"></a>Tipo

*   String

##### <a name="requirements"></a>Requisitos

|Requisito| Valor|
|---|---|
|[Versão do conjunto de requisitos mínimos da caixa de correio](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| 1.0|
|[Nível de permissão mínimo](/outlook/add-ins/understanding-outlook-add-in-permissions)| ReadItem|
|[Modo do Outlook aplicável](/outlook/add-ins/#extension-points)| Escrever ou Ler|

---
---

#### <a name="resturl-string"></a>restUrl: cadeia de caracteres

Obtém a URL do ponto de extremidade de REST para esta conta de email.

O valor `restUrl` pode ser usado para fazer chamadas da [API REST](/outlook/rest/) para a caixa de correio do usuário.

Seu aplicativo deve ter a permissão **ReadItem** especificada em seu manifesto para chamar o membro `restUrl` em modo de leitura.

No modo de composição, é preciso chamar o método [`saveAsync`](Office.context.mailbox.item.md#saveasyncoptions-callback) antes de poder usar o membro `restUrl`. Seu aplicativo deve ter permissões **ReadWriteItem** para chamar o método `saveAsync`.

##### <a name="type"></a>Tipo

*   String

##### <a name="requirements"></a>Requisitos

|Requisito| Valor|
|---|---|
|[Versão do conjunto de requisitos mínimos da caixa de correio](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| 1,5 |
|[Nível de permissão mínimo](/outlook/add-ins/understanding-outlook-add-in-permissions)| ReadItem|
|[Modo do Outlook aplicável](/outlook/add-ins/#extension-points)| Escrever ou Ler|

### <a name="methods"></a>Métodos

#### <a name="addhandlerasynceventtype-handler-options-callback"></a>addHandlerAsync(eventType, handler, [options], [callback])

Adiciona um manipulador de eventos a um evento com suporte.

Atualmente, o único tipo de evento compatível é `Office.EventType.ItemChanged`.

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
|[Versão do conjunto de requisitos mínimos da caixa de correio](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| 1,5 |
|[Nível de permissão mínimo](/outlook/add-ins/understanding-outlook-add-in-permissions)| ReadItem |
|[Modo do Outlook aplicável](/outlook/add-ins/#extension-points)| Escrever ou Ler|

##### <a name="example"></a>Exemplo

```javascript
Office.initialize = function (reason) {
  $(document).ready(function () {
    Office.context.mailbox.addHandlerAsync(Office.EventType.ItemChanged, loadNewItem, function (result) {
      if (result.status === Office.AsyncResultStatus.Failed) {
        // Handle error.
      }
    });
  });
};

function loadNewItem(eventArgs) {
  // Load the properties of the newly selected item.
  loadProps(Office.context.mailbox.item);
};
```

---
---

#### <a name="converttoewsiditemid-restversion--string"></a>convertToEwsId(itemId, restVersion) → {String}

Converte uma ID de item formatada para REST no formato EWS.

> [!NOTE]
> Não há suporte para esse método no Outlook no iOS ou no Android.

IDs de itens recuperadas por meio de uma API REST (como a [API do Email do Outlook](/previous-versions/office/office-365-api/api/version-2.0/mail-rest-operations) ou o [Microsoft Graph](https://graph.microsoft.io/)) usam um formato diferente daquele usado pelos Serviços Web do Exchange (EWS). O método `convertToEwsId` converte uma ID formatada como REST para o formato adequado para EWS.

##### <a name="parameters"></a>Parâmetros

|Nome| Tipo| Descrição|
|---|---|---|
|`itemId`| String|Uma ID de item formatada para APIs REST do Outlook|
|`restVersion`| [Office.MailboxEnums.RestVersion](/javascript/api/outlook/office.mailboxenums.restversion?view=outlook-js-1.7)|Um valor que indica a versão da API REST do Outlook usada para recuperar a ID do item.|

##### <a name="requirements"></a>Requisitos

|Requisito| Valor|
|---|---|
|[Versão do conjunto de requisitos mínimos da caixa de correio](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| 1.3|
|[Nível de permissão mínimo](/outlook/add-ins/understanding-outlook-add-in-permissions)| Restrito|
|[Modo do Outlook aplicável](/outlook/add-ins/#extension-points)| Escrever ou Ler|

##### <a name="returns"></a>Retorna:

Tipo: String

##### <a name="example"></a>Exemplo

```javascript
// Get an item's ID from a REST API.
var restId = 'AAMkAGVlOTZjNTM3LW...';

// Treat restId as coming from the v2.0 version of the Outlook Mail API.
var ewsId = Office.context.mailbox.convertToEwsId(restId, Office.MailboxEnums.RestVersion.v2_0);
```

---
---

#### <a name="converttolocalclienttimetimevalue--localclienttimejavascriptapioutlookofficelocalclienttimeviewoutlook-js-17"></a>convertToLocalClientTime(timeValue) → {[LocalClientTime](/javascript/api/outlook/office.LocalClientTime?view=outlook-js-1.7)}

Obtém um dicionário contendo informações de hora em tempo local do cliente.

Um aplicativo de email para o Outlook em uma área de trabalho ou na Web pode usar fusos horários diferentes para datas e horas. O Outlook em uma área de trabalho usa o fuso horário do computador cliente; O Outlook na Web usa o fuso horário definido no centro de administração do Exchange (Eat). Você deve lidar com valores de data e hora para que os valores exibidos na interface do usuário sejam sempre consistentes com o fuso horário que o usuário espera.

Se o aplicativo de email estiver em execução no Outlook em um cliente desktop `convertToLocalClientTime` , o método retornará um objeto Dictionary com os valores definidos para o fuso horário do computador cliente. Se o aplicativo de email estiver em execução no Outlook na Web, `convertToLocalClientTime` o método retornará um objeto Dictionary com os valores definidos para o fuso horário especificado no Eat.

##### <a name="parameters"></a>Parâmetros

|Nome| Tipo| Descrição|
|---|---|---|
|`timeValue`| Date|Um objeto Date|

##### <a name="requirements"></a>Requisitos

|Requisito| Valor|
|---|---|
|[Versão do conjunto de requisitos mínimos da caixa de correio](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| 1.0|
|[Nível de permissão mínimo](/outlook/add-ins/understanding-outlook-add-in-permissions)| ReadItem|
|[Modo do Outlook aplicável](/outlook/add-ins/#extension-points)| Escrever ou Ler|

##### <a name="returns"></a>Retorna:

Tipo: [LocalClientTime](/javascript/api/outlook/office.LocalClientTime?view=outlook-js-1.7)

---
---

#### <a name="converttorestiditemid-restversion--string"></a>convertToRestId(itemId, restVersion) → {String}

Converte uma ID de item formatada para EWS no formato REST.

> [!NOTE]
> Não há suporte para esse método no Outlook no iOS ou no Android.

IDs de itens recuperadas por EWS ou pela propriedade `itemId` usam um formato diferente daquele usado por APIs REST (como a [API do Email do Outlook](/previous-versions/office/office-365-api/api/version-2.0/mail-rest-operations) ou o [Microsoft Graph](https://graph.microsoft.io/)). O método `convertToRestId` converte uma ID formatada como EWS para o formato adequado para REST.

##### <a name="parameters"></a>Parâmetros

|Nome| Tipo| Descrição|
|---|---|---|
|`itemId`| String|Uma ID de item formatada para os Serviços Web do Exchange (EWS)|
|`restVersion`| [Office.MailboxEnums.RestVersion](/javascript/api/outlook/office.mailboxenums.restversion?view=outlook-js-1.7)|Um valor que indica a versão da API REST do Outlook com a qual a ID convertida será usada.|

##### <a name="requirements"></a>Requisitos

|Requisito| Valor|
|---|---|
|[Versão do conjunto de requisitos mínimos da caixa de correio](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| 1.3|
|[Nível de permissão mínimo](/outlook/add-ins/understanding-outlook-add-in-permissions)| Restrito|
|[Modo do Outlook aplicável](/outlook/add-ins/#extension-points)| Escrever ou Ler|

##### <a name="returns"></a>Retorna:

Tipo: String

##### <a name="example"></a>Exemplo

```javascript
// Get the currently selected item's ID.
var ewsId = Office.context.mailbox.item.itemId;

// Convert to a REST ID for the v2.0 version of the Outlook Mail API.
var restId = Office.context.mailbox.convertToRestId(ewsId, Office.MailboxEnums.RestVersion.v2_0);
```

---
---

#### <a name="converttoutcclienttimeinput--date"></a>convertToUtcClientTime(input) → {Date}

Obtém um objeto Date de um dicionário contendo as informações de hora.

O método `convertToUtcClientTime` converte um dicionário que contém uma data e hora locais para um objeto Date com os valores corretos para a data e hora locais.

##### <a name="parameters"></a>Parâmetros

|Nome| Tipo| Descrição|
|---|---|---|
|`input`| [LocalClientTime](/javascript/api/outlook/office.LocalClientTime?view=outlook-js-1.7)|O valor de hora local a converter.|

##### <a name="requirements"></a>Requisitos

|Requisito| Valor|
|---|---|
|[Versão do conjunto de requisitos mínimos da caixa de correio](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| 1.0|
|[Nível de permissão mínimo](/outlook/add-ins/understanding-outlook-add-in-permissions)| ReadItem|
|[Modo do Outlook aplicável](/outlook/add-ins/#extension-points)| Escrever ou Ler|

##### <a name="returns"></a>Retorna:

Um objeto Date com a hora expressa em UTC.

<dl class="param-type">

<dt>Type</dt>

<dd>Date</dd>

</dl>

---
---

#### <a name="displayappointmentformitemid"></a>displayAppointmentForm(itemId)

Exibe um compromisso de calendário existente.

> [!NOTE]
> Não há suporte para esse método no Outlook no iOS ou no Android.

O método `displayAppointmentForm` abre um compromisso de calendário existente em uma nova janela na área de trabalho ou em uma caixa de diálogo em dispositivos móveis.

No Outlook no Mac, você pode usar esse método para exibir um único compromisso que não faz parte de uma série recorrente ou o compromisso mestre de uma série recorrente, mas não é possível exibir uma instância da série. Isso ocorre porque, no Outlook no Mac, você não pode acessar as propriedades (incluindo a ID do item) de instâncias de uma série recorrente.

No Outlook na Web, este método abre o formulário especificado somente se o corpo do formulário for menor ou igual ao número de caracteres 32 KB.

Se o identificador do item especificado não identificar um compromisso existente, um painel em branco abre no dispositivo ou no computador cliente e nenhuma mensagem de erro será exibida.

##### <a name="parameters"></a>Parâmetros

|Nome| Tipo| Descrição|
|---|---|---|
|`itemId`| String|O identificador dos Serviços Web do Exchange (EWS) para um compromisso de calendário existente.|

##### <a name="requirements"></a>Requisitos

|Requisito| Valor|
|---|---|
|[Versão do conjunto de requisitos mínimos da caixa de correio](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| 1.0|
|[Nível de permissão mínimo](/outlook/add-ins/understanding-outlook-add-in-permissions)| ReadItem|
|[Modo do Outlook aplicável](/outlook/add-ins/#extension-points)| Escrever ou Ler|

##### <a name="example"></a>Exemplo

```javascript
Office.context.mailbox.displayAppointmentForm(appointmentId);
```

---
---

#### <a name="displaymessageformitemid"></a>displayMessageForm(itemId)

Exibe uma mensagem existente.

> [!NOTE]
> Não há suporte para esse método no Outlook no iOS ou no Android.

O método `displayMessageForm` abre uma mensagem existente em uma nova janela na área de trabalho ou em uma caixa de diálogo em dispositivos móveis.

No Outlook na Web, este método abre o formulário especificado somente se o corpo do formulário for menor ou igual ao número de caracteres de 32 KB.

Se o identificador do item especificado não identificar uma mensagem existente, não será exibida mensagem no computador cliente e nenhuma mensagem de erro será retornada.

Não use o método `displayMessageForm` com um `itemId` que representa um compromisso. Use o método `displayAppointmentForm` para exibir um compromisso existente e `displayNewAppointmentForm` para exibir um formulário e criar um novo compromisso.

##### <a name="parameters"></a>Parâmetros

|Nome| Tipo| Descrição|
|---|---|---|
|`itemId`| String|O identificador dos Serviços Web do Exchange (EWS) para uma mensagem existente.|

##### <a name="requirements"></a>Requisitos

|Requisito| Valor|
|---|---|
|[Versão do conjunto de requisitos mínimos da caixa de correio](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| 1.0|
|[Nível de permissão mínimo](/outlook/add-ins/understanding-outlook-add-in-permissions)| ReadItem|
|[Modo Aplicável do Outlook](/outlook/add-ins/#extension-points)| Escrever ou Ler|

##### <a name="example"></a>Exemplo

```javascript
Office.context.mailbox.displayMessageForm(messageId);
```

---
---

#### <a name="displaynewappointmentformparameters"></a>displayNewAppointmentForm(parameters)

Exibe um formulário para criar um compromisso no calendário.

> [!NOTE]
> Não há suporte para esse método no Outlook no iOS ou no Android.

O método `displayNewAppointmentForm` abre um formulário que permite ao usuário criar um novo compromisso ou reunião. Se os parâmetros forem especificados, os campos de formulário do compromisso serão preenchidos automaticamente com o conteúdo dos parâmetros.

No Outlook na Web e dispositivos móveis, este método sempre exibe um formulário com um campo participantes. Se você não especificar quaisquer participantes como argumentos de entrada, o método exibe um formulário com um botão **Salvar**. Se você especificar participantes, o formulário inclui os participantes e um botão **Enviar**.

No cliente avançado do Outlook e no Outlook RT, se você especificar quaisquer participantes ou recursos nos parâmetros `requiredAttendees`, `optionalAttendees`ou `resources`, este método exibirá um formulário de reunião com um botão **Enviar**. Se você não especificar destinatários, este método exibirá um formulário de compromisso com um botão **Salvar e Fechar**.

Se qualquer dos parâmetros exceder os limites de tamanho especificados, ou se um nome de parâmetro desconhecido for especificado, ocorre uma exceção.

##### <a name="parameters"></a>Parâmetros

> [!NOTE]
> Todos os parâmetros são opcionais.

|Nome| Tipo| Descrição|
|---|---|---|
| `parameters` | Object | Um dicionário de parâmetros que descreve o novo compromisso. |
| `parameters.requiredAttendees` | Array.&lt;String&gt; &#124; Array.&lt;[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.7)&gt; | Uma matriz de cadeias de caracteres que contém os endereços de email ou uma matriz contendo um objeto `EmailAddressDetails` para cada um dos participantes obrigatórios do compromisso. A matriz está limitada a um máximo de 100 entradas. |
| `parameters.optionalAttendees` | Array.&lt;String&gt; &#124; Array.&lt;[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.7)&gt; | Uma matriz de cadeias de caracteres que contém os endereços de email ou uma matriz contendo um objeto `EmailAddressDetails` para cada um dos participantes opcionais do compromisso. A matriz está limitada a um máximo de 100 entradas. |
| `parameters.start` | Data | Um objeto `Date` que especifica a data e a hora de início do compromisso. |
| `parameters.end` | Data | Um objeto `Date` que especifica a data e a hora de término do compromisso. |
| `parameters.location` | String | Uma cadeia de caracteres que contém o local do compromisso. A cadeia de caracteres está limitada a um máximo de 255 caracteres. |
| `parameters.resources` | Array.&lt;String&gt; | Uma matriz de cadeias de caracteres que contém os recursos necessários para o compromisso. A matriz está limitada a um máximo de 100 entradas. |
| `parameters.subject` | String | Uma cadeia de caracteres que contém o assunto do compromisso. A cadeia de caracteres está limitada a um máximo de 255 caracteres. |
| `parameters.body` | String | O corpo do compromisso. O conteúdo do corpo está limitado a um tamanho máximo de 32 KB. |

##### <a name="requirements"></a>Requisitos

|Requisito| Valor|
|---|---|
|[Versão do conjunto de requisitos mínimos da caixa de correio](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| 1.0|
|[Nível de permissão mínimo](/outlook/add-ins/understanding-outlook-add-in-permissions)| ReadItem|
|[Modo do Outlook aplicável](/outlook/add-ins/#extension-points)| Read|

##### <a name="example"></a>Exemplo

```javascript
var start = new Date();
var end = new Date();
end.setHours(start.getHours() + 1);

Office.context.mailbox.displayNewAppointmentForm(
  {
    requiredAttendees: ['bob@contoso.com'],
    optionalAttendees: ['sam@contoso.com'],
    start: start,
    end: end,
    location: 'Home',
    resources: ['projector@contoso.com'],
    subject: 'meeting',
    body: 'Hello World!'
  });
```

---
---

#### <a name="displaynewmessageformparameters"></a>displayNewMessageForm (parâmetros)

Exibe um formulário para criar uma nova mensagem.

O `displayNewMessageForm` método abre um formulário que permite ao usuário criar uma nova mensagem. Se os parâmetros forem especificados, os campos de formulário da mensagem serão preenchidos automaticamente com o conteúdo dos parâmetros.

Se qualquer dos parâmetros exceder os limites de tamanho especificados, ou se um nome de parâmetro desconhecido for especificado, ocorre uma exceção.

##### <a name="parameters"></a>Parâmetros

> [!NOTE]
> Todos os parâmetros são opcionais.

|Nome| Tipo| Descrição|
|---|---|---|
| `parameters` | Objeto | Um dicionário de parâmetros que descreve a nova mensagem. |
| `parameters.toRecipients` | Array.&lt;String&gt; &#124; Array.&lt;[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.7)&gt; | Uma matriz de cadeias de caracteres que contém os endereços de email `EmailAddressDetails` ou uma matriz que contém um objeto para cada um dos destinatários na linha para. A matriz está limitada a um máximo de 100 entradas. |
| `parameters.ccRecipients` | Array.&lt;String&gt; &#124; Array.&lt;[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.7)&gt; | Uma matriz de cadeias de caracteres que contém os endereços de email `EmailAddressDetails` ou uma matriz que contém um objeto para cada um dos destinatários na linha CC. A matriz está limitada a um máximo de 100 entradas. |
| `parameters.bccRecipients` | Array.&lt;String&gt; &#124; Array.&lt;[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.7)&gt; | Uma matriz de cadeias de caracteres que contém os endereços de email `EmailAddressDetails` ou uma matriz que contém um objeto para cada um dos destinatários na linha Cco. A matriz está limitada a um máximo de 100 entradas. |
| `parameters.subject` | String | Uma cadeia de caracteres que contém o assunto da mensagem. A cadeia de caracteres está limitada a um máximo de 255 caracteres. |
| `parameters.htmlBody` | String | O corpo HTML da mensagem. O conteúdo do corpo está limitado a um tamanho máximo de 32 KB. |
| `parameters.attachments` | Array.&lt;Object&gt; | Uma matriz de objetos JSON que são anexos de arquivo ou item. |
| `parameters.attachments.type` | String | Indica o tipo de anexo. Deve ser `file` para um anexo de arquivo ou `item` para um anexo de item. |
| `parameters.attachments.name` | String | Uma cadeia de caracteres que contém o nome do anexo, até 255 caracteres de comprimento.|
| `parameters.attachments.url` | String | Usado somente se `type` estiver definido como `file`. O URI do local para o arquivo. |
| `parameters.attachments.isInline` | Booliano | Usado somente se `type` estiver definido como `file`. Se for `true`, indicará que o anexo será mostrado embutido no corpo da mensagem e não deverá ser exibido na lista de anexos. |
| `parameters.attachments.itemId` | Cadeia de caracteres | Usado somente se `type` estiver definido como `item`. A ID do item do EWS do email existente que você deseja anexar à nova mensagem. Isso é uma cadeia de até 100 caracteres. |


##### <a name="requirements"></a>Requisitos

|Requisito| Valor|
|---|---|
|[Versão do conjunto de requisitos mínimos da caixa de correio](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| 1.6 |
|[Nível de permissão mínimo](/outlook/add-ins/understanding-outlook-add-in-permissions)| ReadItem|
|[Modo do Outlook aplicável](/outlook/add-ins/#extension-points)| Read|

##### <a name="example"></a>Exemplo

```javascript
Office.context.mailbox.displayNewMessageForm(
  {
    // Copy the To line from current item.
    toRecipients: Office.context.mailbox.item.to
    ccRecipients: ['sam@contoso.com'],
    subject: 'Outlook add-ins are cool!',
    htmlBody: 'Hello <b>World</b>!<br/><img src="cid:image.png"></i>',
    attachments: [
      {
        type: 'file',
        name: 'image.png',
        url: 'http://contoso.com/image.png',
        isInline: true
      }
    ]
  });
```

---
---

#### <a name="getcallbacktokenasyncoptions-callback"></a>getCallbackTokenAsync([options], callback)

Obtém uma cadeia de caracteres que contém um token usado para chamar APIs REST ou Serviços Web do Exchange.

O método `getCallbackTokenAsync` faz uma chamada assíncrona para obter um token opaco do Exchange Server que hospeda a caixa de correio do usuário. A vida útil do token de retorno de chamada é de 5 minutos.

> [!NOTE]
> É recomendável que suplementos usem as APIs REST em vez dos Serviços Web do Exchange sempre que possível.

**Tokens REST**

Quando um token REST é solicitado (`options.isRest = true`), o token resultante não funcionará para autenticar as chamadas dos Serviços Web do Exchange. O token será limitado em escopo para acesso somente leitura no item atual e seus anexos, a menos que o suplemento tenha especificado a permissão [`ReadWriteMailbox`](/outlook/add-ins/understanding-outlook-add-in-permissions#readwritemailbox-permission) em seu manifesto. Se a permissão `ReadWriteMailbox` tiver sido especificada, o token resultante concederá acesso de leitura/gravação a email, calendário e contatos, incluindo a capacidade de enviar emails.

O suplemento deve usar a propriedade `restUrl` para determinar a URL correta a ser usada ao fazer chamadas da API REST.

**Tokens EWS**

Quando um token EWS é solicitado (`options.isRest = false`), o token resultante não funcionará para autenticar as chamadas de API REST. O token será limitado em escopo para acessar o item atual.

O suplemento deve usar a propriedade `ewsUrl` para determinar a URL correta a ser usada ao fazer chamadas de EWS.

##### <a name="parameters"></a>Parâmetros

|Nome| Tipo| Atributos| Descrição|
|---|---|---|---|
| `options` | Objeto | &lt;opcional&gt; | Um objeto literal que contém uma ou mais das propriedades a seguir. |
| `options.isRest` | Booliano |  &lt;opcional&gt; | Determina se o token fornecido será usado para as APIs REST do Outlook ou Serviços Web do Exchange. O valor padrão é `false`. |
| `options.asyncContext` | Objeto |  &lt;opcional&gt; | Quaisquer dados de estado que são passados ao método assíncrono. |
|`callback`| function||Quando o método for concluído, a função passada ao parâmetro `callback` é chamada com um único parâmetro, `asyncResult`, que é um objeto [`AsyncResult`](/javascript/api/office/office.asyncresult). O token é fornecido como uma cadeia de caracteres na propriedade `asyncResult.value`.|

##### <a name="requirements"></a>Requisitos

|Requisito| Valor|
|---|---|
|[Versão do conjunto de requisitos mínimos da caixa de correio](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| 1,5 |
|[Nível de permissão mínimo](/outlook/add-ins/understanding-outlook-add-in-permissions)| ReadItem|
|[Modo do Outlook aplicável](/outlook/add-ins/#extension-points)| Redação e leitura|

##### <a name="example"></a>Exemplo

```javascript
function getCallbackToken() {
  var options = {
    isRest: true,
    asyncContext: { message: 'Hello World!' }
  };

  Office.context.mailbox.getCallbackTokenAsync(options, cb);
}

function cb(asyncResult) {
  var token = asyncResult.value;
}
```

---
---

#### <a name="getcallbacktokenasynccallback-usercontext"></a>getCallbackTokenAsync(callback, [userContext])

Obtém uma cadeia de caracteres que contém um token usado para obter um anexo ou um item de um Exchange Server.

O método `getCallbackTokenAsync` faz uma chamada assíncrona para obter um token opaco do Exchange Server que hospeda a caixa de correio do usuário. A vida útil do token de retorno de chamada é de 5 minutos.

Você pode passar o token e um identificador de anexo ou um identificador de item a um sistema de terceiros. O sistema de terceiros usa o token como portador da autorização para chamar as operações [GetAttachment](/exchange/client-developer/web-service-reference/getattachment-operation) ou [GetItem](/exchange/client-developer/web-service-reference/getitem-operation) dos Serviços Web do Exchange (EWS) para retornar um anexo ou item. Por exemplo, você pode criar um serviço remoto para [obter anexos do item selecionado](/outlook/add-ins/get-attachments-of-an-outlook-item).

Seu aplicativo deve ter a permissão **ReadItem** especificada em seu manifesto para chamar o método `getCallbackTokenAsync` em modo de leitura.

No modo de composição, você deve chamar o método [`saveAsync`](Office.context.mailbox.item.md#saveasyncoptions-callback) para obter um identificador de item para passar ao método `getCallbackTokenAsync`. Seu aplicativo deve ter permissões **ReadWriteItem** para chamar o método `saveAsync`.

##### <a name="parameters"></a>Parâmetros

|Nome| Tipo| Atributos| Descrição|
|---|---|---|---|
|`callback`| function||Quando o método for concluído, a função passada ao parâmetro `callback` é chamada com um único parâmetro, `asyncResult`, que é um objeto [`AsyncResult`](/javascript/api/office/office.asyncresult). O token é fornecido como uma cadeia de caracteres na propriedade `asyncResult.value`.|
|`userContext`| Objeto| &lt;opcional&gt;|Quaisquer dados de estado que são passados ao método assíncrono.|

##### <a name="requirements"></a>Requisitos

|Requisito| Valor|
|---|---|
|[Versão do conjunto de requisitos mínimos da caixa de correio](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| 1.3|
|[Nível de permissão mínimo](/outlook/add-ins/understanding-outlook-add-in-permissions)| ReadItem|
|[Modo do Outlook aplicável](/outlook/add-ins/#extension-points)| Redação e leitura|

##### <a name="example"></a>Exemplo

```javascript
function getCallbackToken() {
  Office.context.mailbox.getCallbackTokenAsync(cb);
}

function cb(asyncResult) {
  var token = asyncResult.value;
}
```

---
---

#### <a name="getuseridentitytokenasynccallback-usercontext"></a>getUserIdentityTokenAsync(callback, [userContext])

Obtém um símbolo que identifica o usuário e o suplemento do Office.

O método `getUserIdentityTokenAsync` retorna um token que pode ser utilizado para identificar e [autenticar o suplemento e o usuário com um sistema de terceiros](/outlook/add-ins/authentication).

##### <a name="parameters"></a>Parâmetros

|Nome| Tipo| Atributos| Descrição|
|---|---|---|---|
|`callback`| function||Quando o método for concluído, a função passada ao parâmetro `callback` é chamada com um único parâmetro, `asyncResult`, que é um objeto [`AsyncResult`](/javascript/api/office/office.asyncresult).<br/><br/>O token é fornecido como uma cadeia de caracteres na propriedade `asyncResult.value`.|
|`userContext`| Object| &lt;opcional&gt;|Quaisquer dados de estado que são passados ao método assíncrono.|

##### <a name="requirements"></a>Requisitos

|Requisito| Valor|
|---|---|
|[Versão do conjunto de requisitos mínimos da caixa de correio](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| 1.0|
|[Nível de permissão mínimo](/outlook/add-ins/understanding-outlook-add-in-permissions)| ReadItem|
|[Modo Aplicável do Outlook](/outlook/add-ins/#extension-points)| Escrever ou Ler|

##### <a name="example"></a>Exemplo

```javascript
function getIdentityToken() {
  Office.context.mailbox.getUserIdentityTokenAsync(cb);
}

function cb(asyncResult) {
  var token = asyncResult.value;
}
```

---
---

#### <a name="makeewsrequestasyncdata-callback-usercontext"></a>makeEwsRequestAsync(data, callback, [userContext])

Faz uma solicitação assíncrona em um serviço dos EWS (Serviços Web do Exchange) no servidor Exchange que hospeda a caixa de correio do usuário.

> [!NOTE]
> Esse método não tem suporte nas seguintes situações.
> - No Outlook no iOS ou no Android
> - Quando o suplemento é carregado em uma caixa de correio do Gmail
> 
> Nesses casos, os suplementos devem [usar as APIs REST](/outlook/add-ins/use-rest-api) para acessar a caixa de correio do usuário.

O método `makeEwsRequestAsync` envia uma solicitação do EWS em nome do suplemento ao Exchange. Consulte [Chamar serviços Web de um suplemento do Outlook](/outlook/add-ins/web-services#ews-operations-that-add-ins-support) para obter uma lista das operações de EWS compatíveis.

Não é possível solicitar os itens associados da pasta com o método `makeEwsRequestAsync`.

A solicitação XML deve especificar a codificação UTF-8.

```xml
<?xml version="1.0" encoding="utf-8"?>
```

O suplemento deve ter a permissão **ReadWriteMailbox** para usar o método `makeEwsRequestAsync`. Para saber mais sobre como usar a permissão **ReadWriteMailbox** e as operações do EWS que você pode chamar com o método `makeEwsRequestAsync`, confira [Especificar permissões para acesso de suplemento de email na caixa de correio do usuário](/outlook/add-ins/understanding-outlook-add-in-permissions).

> [!NOTE]
> O administrador do servidor deve definir `OAuthAuthentication` como true no diretório do EWS para o Servidor de Acesso para Cliente a fim de habilitar o método `makeEwsRequestAsync` a realizar solicitações do EWS.

##### <a name="version-differences"></a>Diferenças de versão

Ao usar o método `makeEwsRequestAsync` nos aplicativos de email em execução em versões do Outlook anteriores à 15.0.4535.1004, é preciso definir o valor de codificação como `ISO-8859-1`.

```xml
<?xml version="1.0" encoding="iso-8859-1"?>
```

Não é necessário definir o valor de codificação quando o aplicativo de email estiver em execução no Outlook na Web. Você pode determinar se o aplicativo de email está em execução no Outlook ou no Outlook na Web usando a propriedade mailbox.diagnostics.hostName. Você pode determinar que versão do Outlook está em execução usando a propriedade mailbox.diagnostics.hostVersion.

##### <a name="parameters"></a>Parâmetros

|Nome| Tipo| Atributos| Descrição|
|---|---|---|---|
|`data`| String||A solicitação do EWS.|
|`callback`| function||Quando o método for concluído, a função passada ao parâmetro `callback` é chamada com um único parâmetro, `asyncResult`, que é um objeto [`AsyncResult`](/javascript/api/office/office.asyncresult).<br/><br/>O resultado XML da chamada do EWS é fornecido como uma cadeia de caracteres na propriedade `asyncResult.value`. Se o resultado exceder 1 MB de tamanho, será exibida uma mensagem de erro.|
|`userContext`| Objeto| &lt;opcional&gt;|Quaisquer dados de estado que são passados ao método assíncrono.|

##### <a name="requirements"></a>Requisitos

|Requisito| Valor|
|---|---|
|[Versão do conjunto de requisitos mínimos da caixa de correio](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| 1.0|
|[Nível de permissão mínimo](/outlook/add-ins/understanding-outlook-add-in-permissions)| ReadWriteMailbox|
|[Modo do Outlook aplicável](/outlook/add-ins/#extension-points)| Escrever ou Ler|

##### <a name="example"></a>Exemplo

O exemplo a seguir chama `makeEwsRequestAsync` para usar a operação `GetItem` para obter o assunto de um item.

```javascript
function getSubjectRequest(id) {
  // Return a GetItem operation request for the subject of the specified item.
  var request =
    '<?xml version="1.0" encoding="utf-8"?>' +
    '<soap:Envelope xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance"' +
    '               xmlns:xsd="http://www.w3.org/2001/XMLSchema"' +
    '               xmlns:soap="http://schemas.xmlsoap.org/soap/envelope/"' +
    '               xmlns:t="http://schemas.microsoft.com/exchange/services/2006/types">' +
    '  <soap:Header>' +
    '    <RequestServerVersion Version="Exchange2013" xmlns="http://schemas.microsoft.com/exchange/services/2006/types" soap:mustUnderstand="0" />' +
    '  </soap:Header>' +
    '  <soap:Body>' +
    '    <GetItem xmlns="http://schemas.microsoft.com/exchange/services/2006/messages">' +
    '      <ItemShape>' +
    '        <t:BaseShape>IdOnly</t:BaseShape>' +
    '        <t:AdditionalProperties>' +
    '            <t:FieldURI FieldURI="item:Subject"/>' +
    '        </t:AdditionalProperties>' +
    '      </ItemShape>' +
    '      <ItemIds><t:ItemId Id="' + id + '"/></ItemIds>' +
    '    </GetItem>' +
    '  </soap:Body>' +
    '</soap:Envelope>';
  
  return request;
}

function sendRequest() {
  // Create a local variable that contains the mailbox.
  Office.context.mailbox.makeEwsRequestAsync(
    getSubjectRequest(mailbox.item.itemId), callback);
}

function callback(asyncResult)  {
  var result = asyncResult.value;
  var context = asyncResult.asyncContext;

  // Process the returned response here.
}
```

---
---

#### <a name="removehandlerasynceventtype-options-callback"></a>removeHandlerAsync(eventType, handler, [options], [callback])

Remove um manipulador de eventos para um tipo de evento com suporte.

Atualmente, o único tipo de evento compatível é `Office.EventType.ItemChanged`.

##### <a name="parameters"></a>Parâmetros

| Nome | Tipo | Atributos | Descrição |
|---|---|---|---|
| `eventType` | [Office.EventType](office.md#eventtype-string) || O evento que deve revogar o manipulador. |
| `options` | Objeto | &lt;opcional&gt; | Um objeto literal que contém uma ou mais das propriedades a seguir. |
| `options.asyncContext` | Objeto | &lt;opcional&gt; | Os desenvolvedores podem fornecer qualquer objeto que desejarem acessar no método de retorno de chamada. |
| `callback` | function| &lt;opcional&gt;|Quando o método for concluído, a função passada ao parâmetro `callback` é chamada com um único parâmetro, `asyncResult`, que é um objeto [`AsyncResult`](/javascript/api/office/office.asyncresult).|

##### <a name="requirements"></a>Requisitos

|Requisito| Valor|
|---|---|
|[Versão do conjunto de requisitos mínimos da caixa de correio](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| 1,5 |
|[Nível de permissão mínimo](/outlook/add-ins/understanding-outlook-add-in-permissions)| ReadItem |
|[Modo Aplicável do Outlook](/outlook/add-ins/#extension-points)| Escrever ou Ler|
