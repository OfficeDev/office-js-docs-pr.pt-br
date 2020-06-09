---
title: Use as APIs REST do Outlook de um suplemento do Outlook
description: Saiba como usar APIs REST do Outlook a partir de um suplemento do Outlook para obter um token de acesso.
ms.date: 10/31/2019
localization_priority: Normal
ms.openlocfilehash: 7cd26c26e277d7d5fe93664494eb84b4e94bcc47
ms.sourcegitcommit: be23b68eb661015508797333915b44381dd29bdb
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 06/08/2020
ms.locfileid: "44611614"
---
# <a name="use-the-outlook-rest-apis-from-an-outlook-add-in"></a><span data-ttu-id="e75f4-103">Use as APIs REST do Outlook de um suplemento do Outlook</span><span class="sxs-lookup"><span data-stu-id="e75f4-103">Use the Outlook REST APIs from an Outlook add-in</span></span>

<span data-ttu-id="e75f4-p101">O namespace [Office.context.mailbox.item](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md) fornece acesso a vários dos campos comuns das mensagens e dos compromissos. No entanto, em alguns cenários um suplemento talvez precise acessar os dados que não são expostos pelo namespace. Por exemplo, o suplemento pode depender de propriedades personalizadas definidas por um aplicativo externo ou ela precisa pesquisar na caixa de correio do usuário pelas mensagens do mesmo remetente. Nessas situações, as [APIs REST do Outlook](/outlook/rest/index) é o método recomendado para recuperar as informações.</span><span class="sxs-lookup"><span data-stu-id="e75f4-p101">The [Office.context.mailbox.item](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md) namespace provides access to many of the common fields of messages and appointments. However, in some scenarios an add-in may need to access data that is not exposed by the namespace. For example, the add-in may rely on custom properties set by an outside app, or it needs to search the user's mailbox for messages from the same sender. In these scenarios, the [Outlook REST APIs](/outlook/rest/index) is the recommended method to retrieve the information.</span></span>

## <a name="get-an-access-token"></a><span data-ttu-id="e75f4-108">Obter um token de acesso</span><span class="sxs-lookup"><span data-stu-id="e75f4-108">Get an access token</span></span>

<span data-ttu-id="e75f4-p102">As APIs REST do Outlook exigem um token portador no cabeçalho `Authorization`. Normalmente, os aplicativos usam fluxos do OAuth2 para recuperar um token. No entanto, os suplementos podem recuperar um token sem implementar o OAuth2 usando o novo método [Office.context.mailbox.getCallbackTokenAsync](../reference/objectmodel/preview-requirement-set/office.context.mailbox.md#methods) introduzido no conjunto de requisitos de Caixa de Correio 1.5.</span><span class="sxs-lookup"><span data-stu-id="e75f4-p102">The Outlook REST APIs require a bearer token in the `Authorization` header. Typically apps use OAuth2 flows to retrieve a token. However, add-ins can retrieve a token without implementing OAuth2 by using the new [Office.context.mailbox.getCallbackTokenAsync](../reference/objectmodel/preview-requirement-set/office.context.mailbox.md#methods) method introduced in the Mailbox requirement set 1.5.</span></span>

<span data-ttu-id="e75f4-112">Ao definir a opção `isRest` como `true`, você poderá solicitar um token compatível com APIs REST.</span><span class="sxs-lookup"><span data-stu-id="e75f4-112">By setting the `isRest` option to `true`, you can request a token compatible with the REST APIs.</span></span>

### <a name="add-in-permissions-and-token-scope"></a><span data-ttu-id="e75f4-113">Permissões de suplementos e escopo do token</span><span class="sxs-lookup"><span data-stu-id="e75f4-113">Add-in permissions and token scope</span></span>

<span data-ttu-id="e75f4-p103">É importante levar em consideração o nível de acesso que seu suplemento precisará com as APIs REST. Na maioria dos casos, o token retornado por `getCallbackTokenAsync` fornecerá acesso somente leitura ao item atual. Isso é verdadeiro mesmo que seu suplemento especifique o nível de permissão `ReadWriteItem` em seu manifesto.</span><span class="sxs-lookup"><span data-stu-id="e75f4-p103">It is important to consider what level of access your add-in will need via the REST APIs. In most cases, the token returned by `getCallbackTokenAsync` will provide read-only access to the current item only. This is true even if your add-in specifies the `ReadWriteItem` permission level in its manifest.</span></span>

<span data-ttu-id="e75f4-p104">Se seu suplemento exigirá acesso de gravação para o item atual ou outros itens na caixa de correio do usuário, o suplemento precisará especificar o nível de permissão `ReadWriteMailbox` em seu manifesto. Nesse caso, o token retornado conterá acesso de leitura/gravação às mensagens, aos eventos e aos contatos do usuário.</span><span class="sxs-lookup"><span data-stu-id="e75f4-p104">If your add-in will require write access to the current item or other items in the user's mailbox, your add-in must specify the `ReadWriteMailbox` permission level in its manifest. In this case, the token returned will contain read/write access to the user's messages, events, and contacts.</span></span>

### <a name="example"></a><span data-ttu-id="e75f4-119">Exemplo</span><span class="sxs-lookup"><span data-stu-id="e75f4-119">Example</span></span>

```js
Office.context.mailbox.getCallbackTokenAsync({isRest: true}, function(result){
  if (result.status === "succeeded") {
    var accessToken = result.value;

    // Use the access token.
    getCurrentItem(accessToken);
  } else {
    // Handle the error.
  }
});
```

## <a name="get-the-item-id"></a><span data-ttu-id="e75f4-120">Obter a ID do item</span><span class="sxs-lookup"><span data-stu-id="e75f4-120">Get the item ID</span></span>

<span data-ttu-id="e75f4-121">Para recuperar o item atual pela REST, o suplemento precisará da ID do item, formatada corretamente para REST.</span><span class="sxs-lookup"><span data-stu-id="e75f4-121">To retrieve the current item via REST, your add-in will need the item's ID, properly formatted for REST.</span></span> <span data-ttu-id="e75f4-122">Isto é obtido pela propriedade [Office.context.mailbox.item.itemId](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md#properties), mas algumas verificações devem ser feitas para garantir que seja uma ID formatada para REST.</span><span class="sxs-lookup"><span data-stu-id="e75f4-122">This is obtained from the [Office.context.mailbox.item.itemId](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md#properties) property, but some checks should be made to ensure that it is a REST-formatted ID.</span></span>

- <span data-ttu-id="e75f4-123">No Outlook Mobile, o valor retornado por `Office.context.mailbox.item.itemId` é uma ID formatada para REST e pode ser usado como está.</span><span class="sxs-lookup"><span data-stu-id="e75f4-123">In Outlook Mobile, the value returned by `Office.context.mailbox.item.itemId` is a REST-formatted ID and can be used as-is.</span></span>
- <span data-ttu-id="e75f4-124">Em outros clientes do Outlook, o valor retornado por `Office.context.mailbox.item.itemId` é uma ID formatada para EWS e deve ser convertida usando o método [Office.context.mailbox.convertToRestId](../reference/objectmodel/preview-requirement-set/office.context.mailbox.md#methods).</span><span class="sxs-lookup"><span data-stu-id="e75f4-124">In other Outlook clients, the value returned by `Office.context.mailbox.item.itemId` is an EWS-formatted ID, and must be converted using the [Office.context.mailbox.convertToRestId](../reference/objectmodel/preview-requirement-set/office.context.mailbox.md#methods) method.</span></span>
- <span data-ttu-id="e75f4-125">Também é necessário converter a ID do anexo em uma ID com formato REST para usá-la.</span><span class="sxs-lookup"><span data-stu-id="e75f4-125">Note you must also convert Attachment ID to a REST-formatted ID in order to use it.</span></span> <span data-ttu-id="e75f4-126">As IDs devem ser convertidas porque as IDs EWS podem conter valores não seguros para URL que causarão problemas ao REST.</span><span class="sxs-lookup"><span data-stu-id="e75f4-126">The reason the IDs must be converted is that EWS IDs can contain non-URL safe values which will cause problems for REST.</span></span>

<span data-ttu-id="e75f4-127">O suplemento pode determinar em qual cliente do Outlook ele será carregado verificando a propriedade [Office.context.mailbox.diagnostics.hostName](/javascript/api/outlook/office.diagnostics#hostname).</span><span class="sxs-lookup"><span data-stu-id="e75f4-127">Your add-in can determine which Outlook client it is loaded in by checking the [Office.context.mailbox.diagnostics.hostName](/javascript/api/outlook/office.diagnostics#hostname) property.</span></span>

### <a name="example"></a><span data-ttu-id="e75f4-128">Exemplo</span><span class="sxs-lookup"><span data-stu-id="e75f4-128">Example</span></span>

```js
function getItemRestId() {
  if (Office.context.mailbox.diagnostics.hostName === 'OutlookIOS') {
    // itemId is already REST-formatted.
    return Office.context.mailbox.item.itemId;
  } else {
    // Convert to an item ID for API v2.0.
    return Office.context.mailbox.convertToRestId(
      Office.context.mailbox.item.itemId,
      Office.MailboxEnums.RestVersion.v2_0
    );
  }
}
```

## <a name="get-the-rest-api-url"></a><span data-ttu-id="e75f4-129">Obter a URL da API REST</span><span class="sxs-lookup"><span data-stu-id="e75f4-129">Get the REST API URL</span></span>

<span data-ttu-id="e75f4-p107">A informação final que seu suplemento precisa para chamar a API REST é o nome do host que deve usar para enviar solicitações de API. Estas informações estão na propriedade [Office.context.mailbox.restUrl](../reference/objectmodel/preview-requirement-set/office.context.mailbox.md#properties).</span><span class="sxs-lookup"><span data-stu-id="e75f4-p107">The final piece of information your add-in needs to call the REST API is the hostname it should use to send API requests. This information is in the [Office.context.mailbox.restUrl](../reference/objectmodel/preview-requirement-set/office.context.mailbox.md#properties) property.</span></span>

### <a name="example"></a><span data-ttu-id="e75f4-132">Exemplo</span><span class="sxs-lookup"><span data-stu-id="e75f4-132">Example</span></span>

```js
// Example: https://outlook.office.com
var restHost = Office.context.mailbox.restUrl;
```

## <a name="call-the-api"></a><span data-ttu-id="e75f4-133">Chamar a API</span><span class="sxs-lookup"><span data-stu-id="e75f4-133">Call the API</span></span>

<span data-ttu-id="e75f4-134">Depois que seu suplemento tiver o token de acesso, a ID do item e a URL da API REST, ele poderá passar essas informações para um serviço de back-end que chama a API REST ou pode chamá-la diretamente usando o AJAX.</span><span class="sxs-lookup"><span data-stu-id="e75f4-134">After your add-in has the access token, item ID, and REST API URL, it can either pass that information to a back-end service which calls the REST API, or it can call it directly using AJAX.</span></span> <span data-ttu-id="e75f4-135">O exemplo a seguir chama a API REST do Email do Outlook para obter a mensagem atual.</span><span class="sxs-lookup"><span data-stu-id="e75f4-135">The following example calls the Outlook Mail REST API to get the current message.</span></span>

```js
function getCurrentItem(accessToken) {
  // Get the item's REST ID.
  var itemId = getItemRestId();

  // Construct the REST URL to the current item.
  // Details for formatting the URL can be found at
  // https://docs.microsoft.com/previous-versions/office/office-365-api/api/version-2.0/mail-rest-operations#get-messages.
  var getMessageUrl = Office.context.mailbox.restUrl +
    '/v2.0/me/messages/' + itemId;

  $.ajax({
    url: getMessageUrl,
    dataType: 'json',
    headers: { 'Authorization': 'Bearer ' + accessToken }
  }).done(function(item){
    // Message is passed in `item`.
    var subject = item.Subject;
    ...
  }).fail(function(error){
    // Handle error.
  });
}
```

## <a name="see-also"></a><span data-ttu-id="e75f4-136">Confira também</span><span class="sxs-lookup"><span data-stu-id="e75f4-136">See also</span></span>

- <span data-ttu-id="e75f4-137">Confira um exemplo que chama as APIs REST de um suplemento do Outlook em [command-demo](https://github.com/OfficeDev/outlook-add-in-command-demo) no GitHub.</span><span class="sxs-lookup"><span data-stu-id="e75f4-137">For an example that calls the REST APIs from an Outlook add-in, see [command-demo](https://github.com/OfficeDev/outlook-add-in-command-demo) on GitHub.</span></span>
- <span data-ttu-id="e75f4-138">As APIs REST do Outlook também estão disponíveis por meio do ponto de extremidade do Microsoft Graph, mas existem algumas diferenças importantes, inclusive como o suplemento obtém um token de acesso.</span><span class="sxs-lookup"><span data-stu-id="e75f4-138">Outlook REST APIs are also available through the Microsoft Graph endpoint but there are some key differences, including how your add-in gets an access token.</span></span> <span data-ttu-id="e75f4-139">Saiba mais em [API REST do Outlook pelo Microsoft Graph](/outlook/rest/index#outlook-rest-api-via-microsoft-graph).</span><span class="sxs-lookup"><span data-stu-id="e75f4-139">For more information, see [Outlook REST API via Microsoft Graph](/outlook/rest/index#outlook-rest-api-via-microsoft-graph).</span></span>