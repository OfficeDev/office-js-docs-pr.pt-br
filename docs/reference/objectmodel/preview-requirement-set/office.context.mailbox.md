
# <a name="mailbox"></a>caixa de correio

### [Office](Office.md)[.context](Office.context.md). mailbox

Fornece acesso ao modelo de objeto do suplemento do Outlook para o Microsoft Outlook e o Microsoft Outlook na Web.

##### <a name="requirements"></a>Requisitos

|Requisito| Valor|
|---|---|
|[Versão mínima do conjunto de requisitos de caixa de correio](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| 1.0|
|[Nível de permissão mínimo](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| Restrito|
|[Modo do Outlook aplicável](https://docs.microsoft.com/outlook/add-ins/#extension-points)| Redigir ou ler|

##### <a name="members-and-methods"></a>Membros e métodos

| Membro | Tipo |
|--------|------|
| [ewsUrl](#ewsurl-string) | Membro |
| [restUrl](#resturl-string) | Membro |
| [addHandlerAsync](#addhandlerasynceventtype-handler-options-callback) | Método |
| [convertToEwsId](#converttoewsiditemid-restversion--string) | Método |
| [convertToLocalClientTime](#converttolocalclienttimetimevalue--localclienttimejavascriptapioutlookofficelocalclienttime) | Método |
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

### <a name="namespaces"></a>Namespaces

[diagnostics](Office.context.mailbox.diagnostics.md): Fornece informações de diagnóstico para um suplemento do Outlook.

[item](Office.context.mailbox.item.md): Fornece métodos e propriedades para acessar uma mensagem ou um compromisso em um suplemento do Outlook.

[userProfile](Office.context.mailbox.userProfile.md): Fornece informações sobre o usuário em um suplemento do Outlook.

### <a name="members"></a>Membros

#### <a name="ewsurl-string"></a>ewsUrl :String

Obtém a URL do ponto de extremidade dos Serviços Web do Exchange (EWS) para esta conta de email. Somente modo de leitura.

> [!NOTE]
> Esse membro não é compatível com o Outlook para iOS ou o Outlook para Android.

O valor `ewsUrl` pode ser usado por um serviço remoto para fazer chamadas do EWS à caixa de correio do usuário. Por exemplo, você pode criar um serviço remoto para [obter anexos do item selecionado](https://docs.microsoft.com/outlook/add-ins/get-attachments-of-an-outlook-item).

Seu aplicativo deve ter a permissão **ReadItem** especificada em seu manifesto para chamar o membro `ewsUrl` em modo de leitura.

No modo redigir, é preciso chamar o método [`saveAsync`](Office.context.mailbox.item.md#saveasyncoptions-callback) antes de poder usar o membro `ewsUrl`. Seu aplicativo deve ter permissões **ReadWriteItem** para chamar o método `saveAsync`.

##### <a name="type"></a>Tipo:

*   sequência de caracteres

##### <a name="requirements"></a>Requisitos

|Requisito| Valor|
|---|---|
|[Versão mínima do conjunto de requisitos de caixa de correio](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| 1.0|
|[Nível de permissão mínimo](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| ReadItem|
|[Modo do Outlook aplicável](https://docs.microsoft.com/outlook/add-ins/#extension-points)| Redigir ou ler|

#### <a name="resturl-string"></a>restUrl :String

Obtém a URL do ponto de extremidade REST para esta conta de email.

O valor `restUrl` pode ser usado para fazer chamadas da [API REST](https://docs.microsoft.com/outlook/rest/) para a caixa de correio do usuário.

Seu aplicativo deve ter a permissão **ReadItem** especificada em seu manifesto para chamar o membro `restUrl` em modo de leitura.

No modo redigir, é preciso chamar o método [`saveAsync`](Office.context.mailbox.item.md#saveasyncoptions-callback) antes de poder usar o membro `restUrl`. Seu aplicativo deve ter permissões **ReadWriteItem** para chamar o método `saveAsync`.

##### <a name="type"></a>Tipo:

*   sequência de caracteres

##### <a name="requirements"></a>Requisitos

|Requisito| Valor|
|---|---|
|[Versão mínima do conjunto de requisitos de caixa de correio](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| 1.5 |
|[Nível de permissão mínimo](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| ReadItem|
|[Modo do Outlook aplicável](https://docs.microsoft.com/outlook/add-ins/#extension-points)| Redigir ou ler|

### <a name="methods"></a>Métodos

####  <a name="addhandlerasynceventtype-handler-options-callback"></a>addHandlerAsync(eventType, handler, [options], [callback])

Adiciona um manipulador de eventos a um evento com suporte.

Atualmente, os tipos de evento compatíveis são `Office.EventType.ItemChanged` e `Office.EventType.OfficeThemeChanged`.

##### <a name="parameters"></a>Parâmetros:

| Nome | Tipo | Atributos | Descrição |
|---|---|---|---|
| `eventType` | [Office.EventType](office.md#eventtype-string) || O evento que deve chamar o manipulador. |
| `handler` | Função || A função para manipular o evento. A função deve aceitar um parâmetro exclusivo, que é um literal de objeto. A propriedade `type` no parâmetro corresponderá ao parâmetro `eventType` passado para `addHandlerAsync`. |
| `options` | Objeto | &lt;opcional&gt; | Um literal de objeto que contém uma ou mais das propriedades a seguir. |
| `options.asyncContext` | Objeto | &lt;opcional&gt; | Os desenvolvedores podem fornecer qualquer objeto que desejarem para acessar no método de retorno de chamada. |
| `callback` | function| &lt;opcional&gt;|Quando o método for concluído, a função passada no parâmetro `callback` é chamada com um parâmetro único, `asyncResult`, que é um objeto [`AsyncResult`](/javascript/api/office/office.asyncresult).|

##### <a name="requirements"></a>Requisitos

|Requisito| Valor|
|---|---|
|[Versão mínima do conjunto de requisitos de caixa de correio](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| 1.5 |
|[Nível de permissão mínimo](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| ReadItem |
|[Modo do Outlook aplicável](https://docs.microsoft.com/outlook/add-ins/#extension-points)| Redigir ou ler|

##### <a name="example"></a>Exemplo

```
Office.initialize = function (reason) {
  $(document).ready(function () {
    Office.context.mailbox.addHandlerAsync(Office.EventType.ItemChanged, loadNewItem, function (result) {
      if (result.status === Office.AsyncResultStatus.Failed) {
        // Handle error
      }
    });
  });
};

function loadNewItem(eventArgs) {
  // Load the properties of the newly selected item
  loadProps(Office.context.mailbox.item);
};
```

####  <a name="converttoewsiditemid-restversion--string"></a>convertToEwsId(itemId, restVersion) → {String}

Converte uma ID de item formatada para REST em formato EWS.

> [!NOTE]
> Esse método não é compatível com o Outlook para iOS ou o Outlook para Android.

As IDs de itens recuperadas por meio de uma API REST (como a [API de email do Outlook](https://docs.microsoft.com/previous-versions/office/office-365-api/api/version-2.0/mail-rest-operations) ou o [Microsoft Graph](http://graph.microsoft.io/)) usam um formato diferente daquele usado pelos Serviços Web do Exchange (EWS). O método `convertToEwsId` converte uma ID formatada em REST no formato adequado para o EWS.

##### <a name="parameters"></a>Parâmetros:

|Nome| Tipo| Descrição|
|---|---|---|
|`itemId`| sequência de caracteres|Uma ID de item formatada para APIs REST do Outlook|
|`restVersion`| [Office.MailboxEnums.RestVersion](/javascript/api/outlook/office.mailboxenums.restversion)|Um valor que indica a versão da API REST do Outlook usada para recuperar a ID do item.|

##### <a name="requirements"></a>Requisitos

|Requisito| Valor|
|---|---|
|[Versão mínima do conjunto de requisitos de caixa de correio](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| 1.3|
|[Nível de permissão mínimo](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| Restrito|
|[Modo do Outlook aplicável](https://docs.microsoft.com/outlook/add-ins/#extension-points)| Redigir ou ler|

##### <a name="returns"></a>Retorna:

Tipo: String

##### <a name="example"></a>Exemplo

```
// Get an item's ID from a REST API
var restId = 'AAMkAGVlOTZjNTM3LW...';

// Treat restId as coming from the v2.0 version of the
// Outlook Mail API
var ewsId = Office.context.mailbox.convertToEwsId(restId, Office.MailboxEnums.RestVersion.v2_0);
```

####  <a name="converttolocalclienttimetimevalue--localclienttimejavascriptapioutlookofficelocalclienttime"></a>convertToLocalClientTime(timeValue) → {[LocalClientTime](/javascript/api/outlook/office.LocalClientTime)}

Obtém um dicionário contendo informações da hora local do cliente.

As datas e horas usadas por um aplicativo de email para o Outlook ou o Outlook Web App podem usar fusos horários diferentes. O Outlook usa o fuso horário do computador cliente; O Outlook Web App usa o fuso horário definido no Centro de Administração do Exchange (EAC). Você deve manipular valores de data e hora para que os valores exibidos na interface do usuário sejam sempre consistentes com o fuso horário esperado pelo usuário.

Se o aplicativo de email estiver sendo executado no Outlook, o método `convertToLocalClientTime` retornará um objeto de dicionário com os valores definidos para o fuso horário do computador do cliente. Se o aplicativo de email estiver sendo executado no Outlook Web App, o método `convertToLocalClientTime` retornará um objeto de dicionário com os valores definidos para o fuso horário especificado no EAC.

##### <a name="parameters"></a>Parâmetros:

|Nome| Tipo| Descrição|
|---|---|---|
|`timeValue`| Date|Um objeto Date|

##### <a name="requirements"></a>Requisitos

|Requisito| Valor|
|---|---|
|[Versão mínima do conjunto de requisitos de caixa de correio](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| 1.0|
|[Nível de permissão mínimo](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| ReadItem|
|[Modo do Outlook aplicável](https://docs.microsoft.com/outlook/add-ins/#extension-points)| Redigir ou ler|

##### <a name="returns"></a>Retorna:

Tipo: [LocalClientTime](/javascript/api/outlook/office.LocalClientTime)

####  <a name="converttorestiditemid-restversion--string"></a>convertToRestId(itemId, restVersion) → {String}

Converte uma ID de item formatada para EWS em formato REST.

> [!NOTE]
> Esse método não é compatível com o Outlook para iOS ou o Outlook para Android.

As IDs de itens recuperadas por EWS ou pela propriedade `itemId` usam um formato diferente daquele usado por APIs REST (como a [API de email do Outlook](https://docs.microsoft.com/previous-versions/office/office-365-api/api/version-2.0/mail-rest-operations) ou o [Microsoft Graph](http://graph.microsoft.io/)). O método `convertToRestId` converte uma ID formatada como EWS para o formato adequado para REST.

##### <a name="parameters"></a>Parâmetros:

|Nome| Tipo| Descrição|
|---|---|---|
|`itemId`| sequência de caracteres|Uma ID de item formatada para os Serviços Web do Exchange (EWS)|
|`restVersion`| [Office.MailboxEnums.RestVersion](/javascript/api/outlook/office.mailboxenums.restversion)|Um valor indicando a versão da API REST do Outlook com a qual a ID convertida será usada.|

##### <a name="requirements"></a>Requisitos

|Requisito| Valor|
|---|---|
|[Versão mínima do conjunto de requisitos de caixa de correio](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| 1.3|
|[Nível de permissão mínimo](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| Restrito|
|[Modo do Outlook aplicável](https://docs.microsoft.com/outlook/add-ins/#extension-points)| Redigir ou ler|

##### <a name="returns"></a>Retorna:

Tipo: String

##### <a name="example"></a>Exemplo

```
// Get the currently selected item's ID
var ewsId = Office.context.mailbox.item.itemId;

// Convert to a REST ID for the v2.0 version of the
// Outlook Mail API
var restId = Office.context.mailbox.convertToRestId(ewsId, Office.MailboxEnums.RestVersion.v2_0);
```

####  <a name="converttoutcclienttimeinput--date"></a>convertToUtcClientTime(input) → {Date}

Obtém um objeto Date de um dicionário contendo as informações de hora.

O método `convertToUtcClientTime` converte um dicionário que contém uma data e hora locais para um objeto Date com os valores corretos para a data e hora locais.

##### <a name="parameters"></a>Parâmetros:

|Nome| Tipo| Descrição|
|---|---|---|
|`input`| [LocalClientTime](/javascript/api/outlook/office.LocalClientTime)|O valor de hora local a ser convertido.|

##### <a name="requirements"></a>Requisitos

|Requisito| Valor|
|---|---|
|[Versão mínima do conjunto de requisitos de caixa de correio](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| 1.0|
|[Nível de permissão mínimo](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| ReadItem|
|[Modo do Outlook aplicável](https://docs.microsoft.com/outlook/add-ins/#extension-points)| Redigir ou ler|

##### <a name="returns"></a>Retorna:

Um objeto Date com o horário expresso em UTC.

<dl class="param-type">

<dt>Tipo</dt>

<dd>Date</dd>

</dl>

####  <a name="displayappointmentformitemid"></a>displayAppointmentForm(itemId)

Exibe um compromisso de calendário existente.

> [!NOTE]
> Esse método não  é suportado no Outlook para iOS nem no Outlook para Android.

O método `displayAppointmentForm` abre um compromisso de calendário existente em uma nova janela na área de trabalho ou em uma caixa de diálogo em dispositivos móveis.

No Outlook para Mac, você pode usar esse método para exibir um único compromisso que não faz parte de uma série recorrente ou o compromisso principal de uma série recorrente, mas você não pode exibir uma instância da série. Isso ocorre porque no Outlook para Mac você não pode acessar as propriedades (incluindo a ID do item) de instâncias de uma série recorrente.

No aplicativo Web do Outlook, esse método abre o formulário especificado somente se o corpo do formulário for menor ou igual a um número de caracteres de 32KB.

Se o identificador de item especificado não identificar um compromisso existente, um painel em branco será aberto no computador ou no dispositivo cliente e nenhuma mensagem de erro será retornada.

##### <a name="parameters"></a>Parâmetros:

|Nome| Tipo| Descrição|
|---|---|---|
|`itemId`| sequência de caracteres|O identificador de serviços da Web do Exchange (EWS) para um compromisso de calendário existente.|

##### <a name="requirements"></a>Requisitos

|Requisito| Valor|
|---|---|
|[Versão mínima do conjunto de requisitos de caixa de correio](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| 1.0|
|[Nível de permissão mínimo](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| ReadItem|
|[Modo do Outlook aplicável](https://docs.microsoft.com/outlook/add-ins/#extension-points)| Redigir ou ler|

##### <a name="example"></a>Exemplo

```
Office.context.mailbox.displayAppointmentForm(appointmentId);
```

####  <a name="displaymessageformitemid"></a>displayMessageForm(itemId)

Exibe uma mensagem existente.

> [!NOTE]
> Esse método não  é suportado no Outlook para iOS nem no Outlook para Android.

O método `displayMessageForm` abre uma mensagem existente em uma nova janela na área de trabalho ou em uma caixa de diálogo em dispositivos móveis.

No aplicativo Web do Outlook, este método abre o formulário especificado somente se o corpo do formulário for menor que ou igual a um número de caracteres de 32 KB.

Se o identificador do item especificado não identificar uma mensagem existente, não será exibida uma mensagem no computador cliente e nenhuma mensagem de erro será retornada.

Não use o método `displayMessageForm` com um `itemId` que representa um compromisso. Use o método `displayAppointmentForm` para exibir um compromisso existente e `displayNewAppointmentForm` para exibir um formulário e criar um novo compromisso.

##### <a name="parameters"></a>Parâmetros:

|Nome| Tipo| Descrição|
|---|---|---|
|`itemId`| sequência de caracteres|O identificador dos Serviços Web do Exchange (EWS) para uma mensagem existente.|

##### <a name="requirements"></a>Requisitos

|Requisito| Valor|
|---|---|
|[Versão mínima do conjunto de requisitos de caixa de correio](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| 1.0|
|[Nível de permissão mínimo](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| ReadItem|
|[Modo do Outlook aplicável](https://docs.microsoft.com/outlook/add-ins/#extension-points)| Redigir ou ler|

##### <a name="example"></a>Exemplo

```
Office.context.mailbox.displayMessageForm(messageId);
```

#### <a name="displaynewappointmentformparameters"></a>displayNewAppointmentForm(parameters)

Exibe um formulário para criar um novo compromisso no calendário.

> [!NOTE]
> Esse método não  é suportado no Outlook para iOS nem no Outlook para Android.

O método `displayNewAppointmentForm` abre um formulário que permite ao usuário criar um novo compromisso ou reunião. Se os parâmetros forem especificados, os campos de formulário do compromisso serão preenchidos automaticamente com o conteúdo dos parâmetros.

No aplicativo Web do Outlook e no OWA para Dispositivos, esse método sempre exibe um formulário com um campo de participantes. Se você não especificar nenhum participante como argumento de entrada, o método exibirá um formulário com um botão **Salvar** . Se você especificar participantes, o formulário incluirá os participantes e um botão **Enviar**.

No cliente avançado do Outlook e no Outlook RT, se você especificar quaisquer participantes ou recursos nos parâmetros `requiredAttendees`, `optionalAttendees` ou `resources`, o método exibirá um formulário de reunião com um botão **Enviar** . Se você não especificar destinatários, o método exibirá um formulário de compromisso com um botão **Salvar & Fechar**.

Se algum dos parâmetros exceder os limites de tamanho especificados ou se um nome de parâmetro desconhecido for especificado, será gerada uma exceção.

##### <a name="parameters"></a>Parâmetros:

> [!NOTE]
> Todos os parâmetros são opcionais.

|Nome| Tipo| Descrição|
|---|---|---|
| `parameters` | Objeto | Um dicionário de parâmetros que descreve o novo compromisso. |
| `parameters.requiredAttendees` | Array.&lt;String&gt; | Array.&lt;[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)&gt; | Uma matriz de sequências de caracteres que contém os endereços de email ou uma matriz contendo um objeto `EmailAddressDetails` para cada um dos participantes obrigatórios do compromisso. A matriz está limitada a um máximo de 100 entradas. |
| `parameters.optionalAttendees` | Array.&lt;String&gt; | Array.&lt;[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)&gt; | Uma matriz de sequências de caracteres que contém os endereços de email ou uma matriz contendo um objeto `EmailAddressDetails` para cada um dos participantes opcionais do compromisso. A matriz está limitada a um máximo de 100 entradas. |
| `parameters.start` | Date | Um objeto `Date` que especifica a data e a hora de início do compromisso. |
| `parameters.end` | Date | Um objeto `Date` que especifica a data e a hora de término do compromisso. |
| `parameters.location` | sequência de caracteres | Uma sequência de caracteres que contém o local do compromisso. Está limitada a um máximo de 255 caracteres. |
| `parameters.resources` | Array.&lt;String&gt; | Uma matriz de sequências de caracteres que contém os recursos necessários para o compromisso. A matriz está limitada a um máximo de 100 entradas. |
| `parameters.subject` | sequência de caracteres | Uma sequência de caracteres que contém o assunto do compromisso. Está limitada a um máximo de 255 caracteres. |
| `parameters.body` | sequência de caracteres | O corpo do compromisso. O conteúdo do corpo está limitado a um tamanho máximo de 32 KB. |

##### <a name="requirements"></a>Requisitos

|Requisito| Valor|
|---|---|
|[Versão mínima do conjunto de requisitos de caixa de correio](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| 1.0|
|[Nível de permissão mínimo](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| ReadItem|
|[Modo do Outlook aplicável](https://docs.microsoft.com/outlook/add-ins/#extension-points)| Leitura|

##### <a name="example"></a>Exemplo

```
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

#### <a name="displaynewmessageformparameters"></a>displayNewMessageForm(parameters)

Exibe um formulário para criar uma nova mensagem.

O método `displayNewMessageForm` abre um formulário que permite ao usuário criar uma nova mensagem. Quando os parâmetros são especificados, os campos do formulário de mensagem são preenchidos automaticamente com o conteúdo dos parâmetros.

Se algum dos parâmetros exceder os limites de tamanho especificados ou se um nome de parâmetro desconhecido for especificado, uma exceção será gerada.

##### <a name="parameters"></a>Parâmetros:

> [!NOTE]
> Todos os parâmetros são opcionais.

|Nome| Tipo| Descrição|
|---|---|---|
| `parameters` | Objeto | Um dicionário de parâmetros que descreve a nova mensagem. |
| `parameters.toRecipients` | Array.&lt;String&gt; | Array.&lt;[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)&gt; | Uma matriz de sequência de caracteres que contém os endereços de e-mail ou uma matriz que contém um objeto `EmailAddressDetails` para cada um dos destinatários na linha Para. A matriz está limitada a um máximo de 100 entradas. |
| `parameters.ccRecipients` | Array.&lt;String&gt; | Array.&lt;[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)&gt; | Uma matriz de sequência de caracteres que contém os endereços de e-mail ou uma matriz que contém um objeto `EmailAddressDetails` para cada um dos destinatários na linha Cc. A matriz está limitada a um máximo de 100 entradas. |
| `parameters.bccRecipients` | Array.&lt;String&gt; | Array.&lt;[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)&gt; | Uma matriz de sequência de caracteres que contém os endereços de e-mail ou uma matriz que contém um objeto `EmailAddressDetails` para cada um dos destinatários na linha Cco. A matriz está limitada a um máximo de 100 entradas. |
| `parameters.subject` | sequência de caracteres | Uma sequência de caracteres que contém o assunto da mensagem. A sequência de caracteres está limitada a um máximo de 255 caracteres. |
| `parameters.htmlBody` | sequência de caracteres | O corpo HTML da mensagem. O conteúdo do corpo está limitado a um tamanho máximo de 32 KB. |
| `parameters.attachments` | Array.&lt;Object&gt; | Uma matriz de objetos JSON que são anexos de arquivo ou item. |
| `parameters.attachments.type` | sequência de caracteres | Indica o tipo de anexo. Deve ser `file` para um anexo de arquivo ou `item` para um anexo de item. |
| `parameters.attachments.name` | sequência de caracteres | Uma sequência de caracteres que contém o nome do anexo, com até 255 caracteres de comprimento.|
| `parameters.attachments.url` | sequência de caracteres | Usado somente se `type` estiver definido como `file`. A URI do local para o arquivo. |
| `parameters.attachments.isInline` | Booleano | Usado somente se `type` estiver definido como `file`. Se for `true`, indicará que o anexo será mostrado em linha no corpo da mensagem e não deverá ser exibido na lista de anexos. |
| `parameters.attachments.itemId` | sequência de caracteres | Usado somente se `type` estiver definido para `item`. A ID do item do EWS do e-mail existente que deseja anexar na nova mensagem. É uma sequência de caracteres com até 100 caracteres. |


##### <a name="requirements"></a>Requisitos

|Requisito| Valor|
|---|---|
|[Versão mínima do conjunto de requisitos de caixa de correio](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| 1.6 |
|[Nível de permissão mínimo](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| ReadItem|
|[Modo do Outlook aplicável](https://docs.microsoft.com/outlook/add-ins/#extension-points)| Leitura|

##### <a name="example"></a>Exemplo

```
Office.context.mailbox.displayNewMessageForm(
  {
    toRecipients: Office.context.mailbox.item.to, // Copy the To line from current item
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

#### <a name="getcallbacktokenasyncoptions-callback"></a>getCallbackTokenAsync([options], callback)

Obtém uma sequência de caracteres que contém um token usado para chamar APIs REST ou Serviços Web do Exchange.

O método `getCallbackTokenAsync` faz uma chamada assíncrona para obter um token opaco do Exchange Server que hospeda a caixa de correio do usuário. A vida útil do token de retorno de chamada é de 5 minutos.

> [!NOTE]
> É recomendável que suplementos usem as APIs REST em vez de Serviços Web do Exchange sempre que possível. 

**Tokens REST**

Quando um token REST é solicitado (`options.isRest = true`), o token resultante não funcionará para autenticar as chamadas dos Serviços Web do Exchange. O token será limitado em escopo para acesso somente leitura no item atual e seus anexos, a menos que o suplemento tenha especificado a permissão [`ReadWriteMailbox`](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions#readwritemailbox-permission) em seu manifesto. Se a permissão `ReadWriteMailbox` tiver sido especificada, o token resultante concederá acesso de leitura/gravação a e-mail, calendário e contatos, incluindo a capacidade de enviar e-mails.

O suplemento deve usar a propriedade `restUrl` para determinar a URL correta a ser usada ao fazer chamadas da API REST.

**Tokens EWS**

Quando um token EWS é solicitado (`options.isRest = false`), o token resultante não funcionará para autenticar as chamadas de API REST. O token será limitado em escopo para acessar o item atual.

O suplemento deve usar a propriedade `ewsUrl` para determinar a URL correta a ser usada ao fazer chamadas de EWS.

##### <a name="parameters"></a>Parâmetros:

|Nome| Tipo| Atributos| Descrição|
|---|---|---|---|
| `options` | Objeto | &lt;opcional&gt; | Um literal de objeto que contém uma ou mais das propriedades a seguir. |
| `options.isRest` | Booleano |  &lt;opcional&gt; | Determina se o token fornecido será usado para as APIs REST do Outlook ou Serviços Web do Exchange. O valor padrão é `false`. |
| `options.asyncContext` | Objeto |  &lt;opcional&gt; | Quaisquer dados de estado que são passados ao método assíncrono. |
|`callback`| function||Quando o método for concluído, a função passada no parâmetro `callback` é chamada com um parâmetro único, `asyncResult`, que é um objeto [`AsyncResult`](/javascript/api/office/office.asyncresult). O token é fornecido como uma sequência de caracteres na propriedade `asyncResult.value`.|

##### <a name="requirements"></a>Requisitos

|Requisito| Valor|
|---|---|
|[Versão mínima do conjunto de requisitos de caixa de correio](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| 1.5 |
|[Nível de permissão mínimo](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| ReadItem|
|[Modo do Outlook aplicável](https://docs.microsoft.com/outlook/add-ins/#extension-points)| Redigir e ler|

##### <a name="example"></a>Exemplo

```js
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

#### <a name="getcallbacktokenasynccallback-usercontext"></a>getCallbackTokenAsync(callback, [userContext])

Obtém uma sequência de caracteres que contém um token usado para obter um anexo ou um item de um Exchange Server.

O método `getCallbackTokenAsync` faz uma chamada assíncrona para obter um token opaco do Exchange Server que hospeda a caixa de correio do usuário. A vida útil do token de retorno de chamada é de 5 minutos.

Você pode passar o token e um identificador de anexo ou um identificador de item a um sistema de terceiros. O sistema de terceiros usa o token como portador da autorização para chamar as operações [GetAttachment](https://docs.microsoft.com/exchange/client-developer/web-service-reference/getattachment-operation) ou [GetItem](https://docs.microsoft.com/exchange/client-developer/web-service-reference/getitem-operation) dos Serviços Web do Exchange (EWS) para retornar um anexo ou item. Por exemplo, você pode criar um serviço remoto para [obter anexos do item selecionado](https://docs.microsoft.com/outlook/add-ins/get-attachments-of-an-outlook-item).

Seu aplicativo deve ter a permissão **ReadItem** especificada em seu manifesto para chamar o método `getCallbackTokenAsync` em modo de leitura.

No modo de composição, você deve chamar o método [`saveAsync`](Office.context.mailbox.item.md#saveasyncoptions-callback) para obter um identificador de item para passar ao método `getCallbackTokenAsync`. Seu aplicativo deve ter permissões **ReadWriteItem** para chamar o método `saveAsync`.

##### <a name="parameters"></a>Parâmetros:

|Nome| Tipo| Atributos| Descrição|
|---|---|---|---|
|`callback`| function||Quando o método for concluído, a função passada no parâmetro `callback` é chamada com um parâmetro único, `asyncResult`, que é um objeto [`AsyncResult`](/javascript/api/office/office.asyncresult). O token é fornecido como uma sequência de caracteres na propriedade `asyncResult.value`.|
|`userContext`| Objeto| &lt;opcional&gt;|Quaisquer dados de estado que são passados ao método assíncrono.|

##### <a name="requirements"></a>Requisitos

|Requisito| Valor|
|---|---|
|[Versão mínima do conjunto de requisitos de caixa de correio](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| 1.3|
|[Nível de permissão mínimo](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| ReadItem|
|[Modo do Outlook aplicável](https://docs.microsoft.com/outlook/add-ins/#extension-points)| Redigir e ler|

##### <a name="example"></a>Exemplo

```js
function getCallbackToken() {
  Office.context.mailbox.getCallbackTokenAsync(cb);
}

function cb(asyncResult) {
  var token = asyncResult.value;
}
```

####  <a name="getuseridentitytokenasynccallback-usercontext"></a>getUserIdentityTokenAsync(callback, [userContext])

Obtém um token que identifica o usuário e o suplemento do Office.

O método `getUserIdentityTokenAsync` retorna um token que pode ser utilizado para identificar e [autenticar o suplemento e o usuário com um sistema de terceiros](https://docs.microsoft.com/outlook/add-ins/authentication).

##### <a name="parameters"></a>Parâmetros:

|Nome| Tipo| Atributos| Descrição|
|---|---|---|---|
|`callback`| function||Quando o método for concluído, a função passada no parâmetro `callback` é chamada com um parâmetro único, `asyncResult`, que é um objeto [`AsyncResult`](/javascript/api/office/office.asyncresult).<br/><br/>O token é fornecido como uma sequência de caracteres na propriedade `asyncResult.value`.|
|`userContext`| Objeto| &lt;opcional&gt;|Quaisquer dados de estado que são passados ao método assíncrono.|

##### <a name="requirements"></a>Requisitos

|Requisito| Valor|
|---|---|
|[Versão mínima do conjunto de requisitos de caixa de correio](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| 1.0|
|[Nível de permissão mínimo](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| ReadItem|
|[Modo do Outlook aplicável](https://docs.microsoft.com/outlook/add-ins/#extension-points)| Redigir ou ler|

##### <a name="example"></a>Exemplo

```js
function getIdentityToken() {
  Office.context.mailbox.getUserIdentityTokenAsync(cb);
}

function cb(asyncResult) {
  var token = asyncResult.value;
}
```

####  <a name="makeewsrequestasyncdata-callback-usercontext"></a>makeEwsRequestAsync(data, callback, [userContext])

Faz uma solicitação assíncrona em um dos Serviços Web do Exchange (EWS) no Exchange Server que hospeda a caixa de correio do usuário.

> [!NOTE]
> Esse método não é suportado nos seguintes cenários.
> - No Outlook para iOS ou no Outlook para Android
> - Quando o suplemento é carregado em uma caixa de correio do Gmail
> 
> Nesses casos, os suplementos devem [usar APIs REST](https://docs.microsoft.com/outlook/add-ins/use-rest-api) para acessar a caixa de correio do usuário.

O método `makeEwsRequestAsync` envia uma solicitação do EWS em nome do suplemento ao Exchange. Para obter uma lista de operações EWS compatíveis, confira [Chamar serviços Web de um suplemento do Outlook](https://docs.microsoft.com/outlook/add-ins/web-services#ews-operations-that-add-ins-support) .

Não é possível solicitar os itens associados à pasta com o método `makeEwsRequestAsync`.

A solicitação XML deve especificar a codificação UTF-8.

```
<?xml version="1.0" encoding="utf-8"?>
```

O suplemento deve ter a permissão **ReadWriteMailbox** para usar o método `makeEwsRequestAsync`. Para obter mais informações sobre como usar a permissão **ReadWriteMailbox** e as operações do EWS que você pode chamar com o método `makeEwsRequestAsync`, confira [Especificar permissões para acesso do suplemento de email na caixa de correio do usuário](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions).

> [!NOTE]
> O administrador do servidor deve definir `OAuthAuthentication` como verdadeiro no diretório EWS do Servidor de Acesso para Cliente para ativar o método `makeEwsRequestAsync` para fazer solicitações do EWS.

##### <a name="version-differences"></a>Diferenças de versão

Ao usar o método `makeEwsRequestAsync` nos aplicativos de email em execução em versões do Outlook anteriores à 15.0.4535.1004, é preciso definir o valor de codificação como `ISO-8859-1`.

```
<?xml version="1.0" encoding="iso-8859-1"?>
```

Você não precisa definir o valor de codificação quando seu aplicativo de email estiver sendo executado no Outlook na Web. Você pode determinar se o seu aplicativo de email está sendo executado no Outlook ou no Outlook na Web usando a propriedade mailbox.diagnostics.hostName. Você pode determinar qual versão do Outlook está sendo executada usando a propriedade mailbox.diagnostics.hostVersion.

##### <a name="parameters"></a>Parâmetros:

|Nome| Tipo| Atributos| Descrição|
|---|---|---|---|
|`data`| sequência de caracteres||A solicitação do EWS.|
|`callback`| function||Quando o método for concluído, a função passada no parâmetro `callback` é chamada com um parâmetro único, `asyncResult`, que é um objeto [`AsyncResult`](/javascript/api/office/office.asyncresult).<br/><br/>O resultado XML da chamada do EWS é fornecido como uma sequência de caracteres na propriedade `asyncResult.value`. Se o resultado exceder 1 MB, será exibida uma mensagem de erro.|
|`userContext`| Objeto| &lt;opcional&gt;|Quaisquer dados de estado que são passados ao método assíncrono.|

##### <a name="requirements"></a>Requisitos

|Requisito| Valor|
|---|---|
|[Versão mínima do conjunto de requisitos de caixa de correio](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| 1.0|
|[Nível de permissão mínimo](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| ReadWriteMailbox|
|[Modo do Outlook aplicável](https://docs.microsoft.com/outlook/add-ins/#extension-points)| Redigir ou ler|

##### <a name="example"></a>Exemplo

O exemplo a seguir chama `makeEwsRequestAsync` para usar a operação `GetItem` para obter o assunto de um item.

```js
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