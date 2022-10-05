---
title: Use as APIs REST do Outlook de um suplemento do Outlook
description: Saiba como usar APIs REST do Outlook a partir de um suplemento do Outlook para obter um token de acesso.
ms.date: 10/03/2022
ms.localizationpriority: medium
ms.openlocfilehash: 9f62b2514f05341531a826c29e18c593a590fca0
ms.sourcegitcommit: 005783ddd43cf6582233be1be6e3463d7ab9b0e5
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 10/05/2022
ms.locfileid: "68467213"
---
# <a name="use-the-outlook-rest-apis-from-an-outlook-add-in"></a>Use as APIs REST do Outlook de um suplemento do Outlook

The [Office.context.mailbox.item](/javascript/api/requirement-sets/outlook/preview-requirement-set/office.context.mailbox.item) namespace provides access to many of the common fields of messages and appointments. However, in some scenarios an add-in may need to access data that is not exposed by the namespace. For example, the add-in may rely on custom properties set by an outside app, or it needs to search the user's mailbox for messages from the same sender. In these scenarios, the [Outlook REST APIs](/outlook/rest) is the recommended method to retrieve the information.

> [!IMPORTANT]
> **As APIs REST do Outlook foram preteridas**
>
> Os pontos de extremidade REST do Outlook serão totalmente desativados em 30 de novembro de 2022 (para obter mais detalhes, consulte o comunicado de [novembro de 2020](https://developer.microsoft.com/graph/blogs/outlook-rest-api-v2-0-deprecation-notice/)). Você deve migrar suplementos existentes para usar o [Microsoft Graph](/outlook/rest#outlook-rest-api-via-microsoft-graph). Para obter diretrizes, consulte [Comparar os pontos de extremidade da API REST do Microsoft Graph e do Outlook](/outlook/rest/compare-graph).
>
> Para ajudá-lo com a migração, os suplementos ativos que usam o serviço REST estão qualificados para uma isenção para continuar usando o serviço até que o suporte estendido termine para o [Outlook 2019 em 14 de outubro de 2025](/lifecycle/end-of-support/end-of-support-2025). Isso inclui novos suplementos desenvolvidos após 30 de novembro de 2022. A isenção é baseada na ID de manifesto do suplemento e se aplica a suplementos hospedados pelo AppSource e liberados de forma privada.
>
> A identificação automática de tráfego de suplementos do Outlook que usam o serviço REST está sendo testada para validação de isenção. Se você quiser participar desta fase de teste, preencha o formulário de verificação do suplemento da [API REST](https://aka.ms/RESTCheck) antes de novembro de 2022. Para obter mais informações, consulte a postagem no blog de chamada da comunidade de [suplementos do Office de agosto de 2022](https://pnp.github.io/blog/office-add-ins-community-call/2022-08-10/).

## <a name="get-an-access-token"></a>Obter um token de acesso

The Outlook REST APIs require a bearer token in the `Authorization` header. Typically apps use OAuth2 flows to retrieve a token. However, add-ins can retrieve a token without implementing OAuth2 by using the new [Office.context.mailbox.getCallbackTokenAsync](/javascript/api/requirement-sets/outlook/preview-requirement-set/office.context.mailbox#methods) method introduced in the Mailbox requirement set 1.5.

Ao definir a opção `isRest` como `true`, você poderá solicitar um token compatível com APIs REST.

### <a name="add-in-permissions-and-token-scope"></a>Permissões de suplementos e escopo do token

É importante levar em consideração o nível de acesso que seu suplemento precisará com as APIs REST. Na maioria dos casos, o token retornado por `getCallbackTokenAsync` fornecerá acesso somente leitura ao item atual. Isso é verdadeiro mesmo se o suplemento especificar o nível de permissão de item de leitura [/](understanding-outlook-add-in-permissions.md#readwrite-item-permission) gravação em seu manifesto.

Se o suplemento exigir acesso de gravação ao item atual ou a outros itens na caixa de correio do usuário, o suplemento deverá especificar a permissão de leitura [/gravação da caixa de correio](understanding-outlook-add-in-permissions.md#readwrite-mailbox-permission).
nível em seu manifesto. Nesse caso, o token retornado conterá acesso de leitura/gravação às mensagens, aos eventos e aos contatos do usuário.

### <a name="example"></a>Exemplo

```js
Office.context.mailbox.getCallbackTokenAsync({isRest: true}, function(result){
  if (result.status === "succeeded") {
    const accessToken = result.value;

    // Use the access token.
    getCurrentItem(accessToken);
  } else {
    // Handle the error.
  }
});
```

## <a name="get-the-item-id"></a>Obter a ID do item

Para recuperar o item atual pela REST, o suplemento precisará da ID do item, formatada corretamente para REST. Isto é obtido pela propriedade [Office.context.mailbox.item.itemId](/javascript/api/requirement-sets/outlook/preview-requirement-set/office.context.mailbox.item#properties), mas algumas verificações devem ser feitas para garantir que seja uma ID formatada para REST.

- No Outlook Mobile, o valor retornado por `Office.context.mailbox.item.itemId` é uma ID formatada para REST e pode ser usado como está.
- Em outros clientes do Outlook, o valor retornado por `Office.context.mailbox.item.itemId` é uma ID formatada para EWS e deve ser convertida usando o método [Office.context.mailbox.convertToRestId](/javascript/api/requirement-sets/outlook/preview-requirement-set/office.context.mailbox#methods).
- Também é necessário converter a ID do anexo em uma ID com formato REST para usá-la. As IDs devem ser convertidas porque as IDs EWS podem conter valores não seguros para URL que causarão problemas ao REST.

O suplemento pode determinar em qual cliente do Outlook ele será carregado verificando a propriedade [Office.context.mailbox.diagnostics.hostName](/javascript/api/outlook/office.diagnostics#outlook-office-diagnostics-hostname-member).

### <a name="example"></a>Exemplo

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

## <a name="get-the-rest-api-url"></a>Obter a URL da API REST

The final piece of information your add-in needs to call the REST API is the hostname it should use to send API requests. This information is in the [Office.context.mailbox.restUrl](/javascript/api/requirement-sets/outlook/preview-requirement-set/office.context.mailbox#properties) property.

### <a name="example"></a>Exemplo

```js
// Example: https://outlook.office.com
const restHost = Office.context.mailbox.restUrl;
```

## <a name="call-the-api"></a>Chamar a API

Depois que seu suplemento tiver o token de acesso, a ID do item e a URL da API REST, ele poderá passar essas informações para um serviço de back-end que chama a API REST ou pode chamá-la diretamente usando o AJAX. O exemplo a seguir chama a API REST do Email do Outlook para obter a mensagem atual.

> [!IMPORTANT]
> Para implantações locais do Exchange, as solicitações do lado do cliente que usam AJAX ou bibliotecas semelhantes falham porque o CORS não tem suporte nessa configuração de servidor.

```js
function getCurrentItem(accessToken) {
  // Get the item's REST ID.
  const itemId = getItemRestId();

  // Construct the REST URL to the current item.
  // Details for formatting the URL can be found at
  // https://learn.microsoft.com/previous-versions/office/office-365-api/api/version-2.0/mail-rest-operations#get-messages.
  const getMessageUrl = Office.context.mailbox.restUrl +
    '/v2.0/me/messages/' + itemId;

  $.ajax({
    url: getMessageUrl,
    dataType: 'json',
    headers: { 'Authorization': 'Bearer ' + accessToken }
  }).done(function(item){
    // Message is passed in `item`.
    const subject = item.Subject;
    ...
  }).fail(function(error){
    // Handle error.
  });
}
```

## <a name="see-also"></a>Confira também

- Confira um exemplo que chama as APIs REST de um suplemento do Outlook em [command-demo](https://github.com/OfficeDev/outlook-add-in-command-demo) no GitHub.
- As APIs REST do Outlook também estão disponíveis por meio do ponto de extremidade do Microsoft Graph, mas existem algumas diferenças importantes, inclusive como o suplemento obtém um token de acesso. Saiba mais em [API REST do Outlook pelo Microsoft Graph](/outlook/rest/index#outlook-rest-api-via-microsoft-graph).
