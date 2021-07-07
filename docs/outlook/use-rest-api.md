---
title: Use as APIs REST do Outlook de um suplemento do Outlook
description: Saiba como usar APIs REST do Outlook a partir de um suplemento do Outlook para obter um token de acesso.
ms.date: 07/06/2021
localization_priority: Normal
ms.openlocfilehash: 9f6642afcfae8efd54c4ade6165aa2a6823e3bd2
ms.sourcegitcommit: 488b26b29c7534e3bbc862b688ed2319cc028f71
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 07/07/2021
ms.locfileid: "53315145"
---
# <a name="use-the-outlook-rest-apis-from-an-outlook-add-in"></a><span data-ttu-id="46f42-103">Use as APIs REST do Outlook de um suplemento do Outlook</span><span class="sxs-lookup"><span data-stu-id="46f42-103">Use the Outlook REST APIs from an Outlook add-in</span></span>

<span data-ttu-id="46f42-p101">O namespace [Office.context.mailbox.item](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md) fornece acesso a vários dos campos comuns das mensagens e dos compromissos. No entanto, em alguns cenários um suplemento talvez precise acessar os dados que não são expostos pelo namespace. Por exemplo, o suplemento pode depender de propriedades personalizadas definidas por um aplicativo externo ou ela precisa pesquisar na caixa de correio do usuário pelas mensagens do mesmo remetente. Nessas situações, as [APIs REST do Outlook](/outlook/rest) é o método recomendado para recuperar as informações.</span><span class="sxs-lookup"><span data-stu-id="46f42-p101">The [Office.context.mailbox.item](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md) namespace provides access to many of the common fields of messages and appointments. However, in some scenarios an add-in may need to access data that is not exposed by the namespace. For example, the add-in may rely on custom properties set by an outside app, or it needs to search the user's mailbox for messages from the same sender. In these scenarios, the [Outlook REST APIs](/outlook/rest) is the recommended method to retrieve the information.</span></span>

> [!IMPORTANT]
> <span data-ttu-id="46f42-108">**As OUTLOOK REST são preteridas**</span><span class="sxs-lookup"><span data-stu-id="46f42-108">**The Outlook REST APIs are deprecated**</span></span>
>
> <span data-ttu-id="46f42-109">Os Outlook de extremidade REST serão totalmente desativados em novembro de 2022 (para obter mais detalhes, consulte o comunicado de novembro de [2020](https://developer.microsoft.com/graph/blogs/outlook-rest-api-v2-0-deprecation-notice/)).</span><span class="sxs-lookup"><span data-stu-id="46f42-109">The Outlook REST endpoints will be fully decommissioned in November 2022 (for more details, refer to the [November 2020 announcement](https://developer.microsoft.com/graph/blogs/outlook-rest-api-v2-0-deprecation-notice/)).</span></span> <span data-ttu-id="46f42-110">Você deve migrar os complementos existentes para usar o [Microsoft Graph](/outlook/rest#outlook-rest-api-via-microsoft-graph).</span><span class="sxs-lookup"><span data-stu-id="46f42-110">You should migrate existing add-ins to use [Microsoft Graph](/outlook/rest#outlook-rest-api-via-microsoft-graph).</span></span> <span data-ttu-id="46f42-111">Além disso, compare os pontos de extremidade Graph e Outlook API REST da [Microsoft.](/outlook/rest/compare-graph)</span><span class="sxs-lookup"><span data-stu-id="46f42-111">Also, [Compare Microsoft Graph and Outlook REST API endpoints](/outlook/rest/compare-graph).</span></span>

## <a name="get-an-access-token"></a><span data-ttu-id="46f42-112">Obter um token de acesso</span><span class="sxs-lookup"><span data-stu-id="46f42-112">Get an access token</span></span>

<span data-ttu-id="46f42-p103">As APIs REST do Outlook exigem um token portador no cabeçalho `Authorization`. Normalmente, os aplicativos usam fluxos do OAuth2 para recuperar um token. No entanto, os suplementos podem recuperar um token sem implementar o OAuth2 usando o novo método [Office.context.mailbox.getCallbackTokenAsync](../reference/objectmodel/preview-requirement-set/office.context.mailbox.md#methods) introduzido no conjunto de requisitos de Caixa de Correio 1.5.</span><span class="sxs-lookup"><span data-stu-id="46f42-p103">The Outlook REST APIs require a bearer token in the `Authorization` header. Typically apps use OAuth2 flows to retrieve a token. However, add-ins can retrieve a token without implementing OAuth2 by using the new [Office.context.mailbox.getCallbackTokenAsync](../reference/objectmodel/preview-requirement-set/office.context.mailbox.md#methods) method introduced in the Mailbox requirement set 1.5.</span></span>

<span data-ttu-id="46f42-116">Ao definir a opção `isRest` como `true`, você poderá solicitar um token compatível com APIs REST.</span><span class="sxs-lookup"><span data-stu-id="46f42-116">By setting the `isRest` option to `true`, you can request a token compatible with the REST APIs.</span></span>

### <a name="add-in-permissions-and-token-scope"></a><span data-ttu-id="46f42-117">Permissões de suplementos e escopo do token</span><span class="sxs-lookup"><span data-stu-id="46f42-117">Add-in permissions and token scope</span></span>

<span data-ttu-id="46f42-p104">É importante levar em consideração o nível de acesso que seu suplemento precisará com as APIs REST. Na maioria dos casos, o token retornado por `getCallbackTokenAsync` fornecerá acesso somente leitura ao item atual. Isso é verdadeiro mesmo que seu suplemento especifique o nível de permissão `ReadWriteItem` em seu manifesto.</span><span class="sxs-lookup"><span data-stu-id="46f42-p104">It is important to consider what level of access your add-in will need via the REST APIs. In most cases, the token returned by `getCallbackTokenAsync` will provide read-only access to the current item only. This is true even if your add-in specifies the `ReadWriteItem` permission level in its manifest.</span></span>

<span data-ttu-id="46f42-p105">Se seu suplemento exigirá acesso de gravação para o item atual ou outros itens na caixa de correio do usuário, o suplemento precisará especificar o nível de permissão `ReadWriteMailbox` em seu manifesto. Nesse caso, o token retornado conterá acesso de leitura/gravação às mensagens, aos eventos e aos contatos do usuário.</span><span class="sxs-lookup"><span data-stu-id="46f42-p105">If your add-in will require write access to the current item or other items in the user's mailbox, your add-in must specify the `ReadWriteMailbox` permission level in its manifest. In this case, the token returned will contain read/write access to the user's messages, events, and contacts.</span></span>

### <a name="example"></a><span data-ttu-id="46f42-123">Exemplo</span><span class="sxs-lookup"><span data-stu-id="46f42-123">Example</span></span>

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

## <a name="get-the-item-id"></a><span data-ttu-id="46f42-124">Obter a ID do item</span><span class="sxs-lookup"><span data-stu-id="46f42-124">Get the item ID</span></span>

<span data-ttu-id="46f42-125">Para recuperar o item atual pela REST, o suplemento precisará da ID do item, formatada corretamente para REST.</span><span class="sxs-lookup"><span data-stu-id="46f42-125">To retrieve the current item via REST, your add-in will need the item's ID, properly formatted for REST.</span></span> <span data-ttu-id="46f42-126">Isto é obtido pela propriedade [Office.context.mailbox.item.itemId](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md#properties), mas algumas verificações devem ser feitas para garantir que seja uma ID formatada para REST.</span><span class="sxs-lookup"><span data-stu-id="46f42-126">This is obtained from the [Office.context.mailbox.item.itemId](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md#properties) property, but some checks should be made to ensure that it is a REST-formatted ID.</span></span>

- <span data-ttu-id="46f42-127">No Outlook Mobile, o valor retornado por `Office.context.mailbox.item.itemId` é uma ID formatada para REST e pode ser usado como está.</span><span class="sxs-lookup"><span data-stu-id="46f42-127">In Outlook Mobile, the value returned by `Office.context.mailbox.item.itemId` is a REST-formatted ID and can be used as-is.</span></span>
- <span data-ttu-id="46f42-128">Em outros clientes do Outlook, o valor retornado por `Office.context.mailbox.item.itemId` é uma ID formatada para EWS e deve ser convertida usando o método [Office.context.mailbox.convertToRestId](../reference/objectmodel/preview-requirement-set/office.context.mailbox.md#methods).</span><span class="sxs-lookup"><span data-stu-id="46f42-128">In other Outlook clients, the value returned by `Office.context.mailbox.item.itemId` is an EWS-formatted ID, and must be converted using the [Office.context.mailbox.convertToRestId](../reference/objectmodel/preview-requirement-set/office.context.mailbox.md#methods) method.</span></span>
- <span data-ttu-id="46f42-129">Também é necessário converter a ID do anexo em uma ID com formato REST para usá-la.</span><span class="sxs-lookup"><span data-stu-id="46f42-129">Note you must also convert Attachment ID to a REST-formatted ID in order to use it.</span></span> <span data-ttu-id="46f42-130">As IDs devem ser convertidas porque as IDs EWS podem conter valores não seguros para URL que causarão problemas ao REST.</span><span class="sxs-lookup"><span data-stu-id="46f42-130">The reason the IDs must be converted is that EWS IDs can contain non-URL safe values which will cause problems for REST.</span></span>

<span data-ttu-id="46f42-131">O suplemento pode determinar em qual cliente do Outlook ele será carregado verificando a propriedade [Office.context.mailbox.diagnostics.hostName](/javascript/api/outlook/office.diagnostics#hostname).</span><span class="sxs-lookup"><span data-stu-id="46f42-131">Your add-in can determine which Outlook client it is loaded in by checking the [Office.context.mailbox.diagnostics.hostName](/javascript/api/outlook/office.diagnostics#hostname) property.</span></span>

### <a name="example"></a><span data-ttu-id="46f42-132">Exemplo</span><span class="sxs-lookup"><span data-stu-id="46f42-132">Example</span></span>

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

## <a name="get-the-rest-api-url"></a><span data-ttu-id="46f42-133">Obter a URL da API REST</span><span class="sxs-lookup"><span data-stu-id="46f42-133">Get the REST API URL</span></span>

<span data-ttu-id="46f42-p108">A informação final que seu suplemento precisa para chamar a API REST é o nome do host que deve usar para enviar solicitações de API. Estas informações estão na propriedade [Office.context.mailbox.restUrl](../reference/objectmodel/preview-requirement-set/office.context.mailbox.md#properties).</span><span class="sxs-lookup"><span data-stu-id="46f42-p108">The final piece of information your add-in needs to call the REST API is the hostname it should use to send API requests. This information is in the [Office.context.mailbox.restUrl](../reference/objectmodel/preview-requirement-set/office.context.mailbox.md#properties) property.</span></span>

### <a name="example"></a><span data-ttu-id="46f42-136">Exemplo</span><span class="sxs-lookup"><span data-stu-id="46f42-136">Example</span></span>

```js
// Example: https://outlook.office.com
var restHost = Office.context.mailbox.restUrl;
```

## <a name="call-the-api"></a><span data-ttu-id="46f42-137">Chamar a API</span><span class="sxs-lookup"><span data-stu-id="46f42-137">Call the API</span></span>

<span data-ttu-id="46f42-138">Depois que seu suplemento tiver o token de acesso, a ID do item e a URL da API REST, ele poderá passar essas informações para um serviço de back-end que chama a API REST ou pode chamá-la diretamente usando o AJAX.</span><span class="sxs-lookup"><span data-stu-id="46f42-138">After your add-in has the access token, item ID, and REST API URL, it can either pass that information to a back-end service which calls the REST API, or it can call it directly using AJAX.</span></span> <span data-ttu-id="46f42-139">O exemplo a seguir chama a API REST do Email do Outlook para obter a mensagem atual.</span><span class="sxs-lookup"><span data-stu-id="46f42-139">The following example calls the Outlook Mail REST API to get the current message.</span></span>

> [!IMPORTANT]
> <span data-ttu-id="46f42-140">Para implantações Exchange locais, as solicitações do lado do cliente usando AJAX ou bibliotecas semelhantes falham porque o CORS não tem suporte nessa configuração de servidor.</span><span class="sxs-lookup"><span data-stu-id="46f42-140">For on-premises Exchange deployments, client-side requests using AJAX or similar libraries fail because CORS isn't supported in that server setup.</span></span>

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

## <a name="see-also"></a><span data-ttu-id="46f42-141">Confira também</span><span class="sxs-lookup"><span data-stu-id="46f42-141">See also</span></span>

- <span data-ttu-id="46f42-142">Confira um exemplo que chama as APIs REST de um suplemento do Outlook em [command-demo](https://github.com/OfficeDev/outlook-add-in-command-demo) no GitHub.</span><span class="sxs-lookup"><span data-stu-id="46f42-142">For an example that calls the REST APIs from an Outlook add-in, see [command-demo](https://github.com/OfficeDev/outlook-add-in-command-demo) on GitHub.</span></span>
- <span data-ttu-id="46f42-143">As APIs REST do Outlook também estão disponíveis por meio do ponto de extremidade do Microsoft Graph, mas existem algumas diferenças importantes, inclusive como o suplemento obtém um token de acesso.</span><span class="sxs-lookup"><span data-stu-id="46f42-143">Outlook REST APIs are also available through the Microsoft Graph endpoint but there are some key differences, including how your add-in gets an access token.</span></span> <span data-ttu-id="46f42-144">Saiba mais em [API REST do Outlook pelo Microsoft Graph](/outlook/rest/index#outlook-rest-api-via-microsoft-graph).</span><span class="sxs-lookup"><span data-stu-id="46f42-144">For more information, see [Outlook REST API via Microsoft Graph](/outlook/rest/index#outlook-rest-api-via-microsoft-graph).</span></span>