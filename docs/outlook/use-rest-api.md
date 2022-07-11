---
title: Use as APIs REST do Outlook de um suplemento do Outlook
description: Saiba como usar APIs REST do Outlook a partir de um suplemento do Outlook para obter um token de acesso.
ms.date: 07/08/2022
ms.localizationpriority: medium
ms.openlocfilehash: 7c02b878b6636e6736ada4a29d123dd8ff772393
ms.sourcegitcommit: d8ea4b761f44d3227b7f2c73e52f0d2233bf22e2
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 07/11/2022
ms.locfileid: "66712962"
---
# <a name="use-the-outlook-rest-apis-from-an-outlook-add-in"></a>Use as APIs REST do Outlook de um suplemento do Outlook

O namespace [Office.context.mailbox.item](/javascript/api/requirement-sets/outlook/preview-requirement-set/office.context.mailbox.item) fornece acesso a vários dos campos comuns das mensagens e dos compromissos. No entanto, em alguns cenários um suplemento talvez precise acessar os dados que não são expostos pelo namespace. Por exemplo, o suplemento pode depender de propriedades personalizadas definidas por um aplicativo externo ou ela precisa pesquisar na caixa de correio do usuário pelas mensagens do mesmo remetente. Nessas situações, as [APIs REST do Outlook](/outlook/rest) é o método recomendado para recuperar as informações.

> [!IMPORTANT]
> **As APIs REST do Outlook foram preteridas**
>
> Os pontos de extremidade REST do Outlook serão totalmente desativados em novembro de 2022 (para obter mais detalhes, consulte o comunicado de [novembro de 2020](https://developer.microsoft.com/graph/blogs/outlook-rest-api-v2-0-deprecation-notice/)). Você deve migrar suplementos existentes para usar o [Microsoft Graph](/outlook/rest#outlook-rest-api-via-microsoft-graph). Além disso, [compare os pontos de extremidade da API REST do Microsoft Graph e do Outlook](/outlook/rest/compare-graph).

## <a name="get-an-access-token"></a>Obter um token de acesso

As APIs REST do Outlook exigem um token portador no cabeçalho `Authorization`. Normalmente, os aplicativos usam fluxos do OAuth2 para recuperar um token. No entanto, os suplementos podem recuperar um token sem implementar o OAuth2 usando o novo método [Office.context.mailbox.getCallbackTokenAsync](/javascript/api/requirement-sets/outlook/preview-requirement-set/office.context.mailbox#methods) introduzido no conjunto de requisitos de Caixa de Correio 1.5.

Ao definir a opção `isRest` como `true`, você poderá solicitar um token compatível com APIs REST.

### <a name="add-in-permissions-and-token-scope"></a>Permissões de suplementos e escopo do token

É importante levar em consideração o nível de acesso que seu suplemento precisará com as APIs REST. Na maioria dos casos, o token retornado por `getCallbackTokenAsync` fornecerá acesso somente leitura ao item atual. Isso é verdadeiro mesmo que seu suplemento especifique o nível de permissão `ReadWriteItem` em seu manifesto.

Se seu suplemento exigirá acesso de gravação para o item atual ou outros itens na caixa de correio do usuário, o suplemento precisará especificar o nível de permissão `ReadWriteMailbox` em seu manifesto. Nesse caso, o token retornado conterá acesso de leitura/gravação às mensagens, aos eventos e aos contatos do usuário.

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

A informação final que seu suplemento precisa para chamar a API REST é o nome do host que deve usar para enviar solicitações de API. Estas informações estão na propriedade [Office.context.mailbox.restUrl](/javascript/api/requirement-sets/outlook/preview-requirement-set/office.context.mailbox#properties).

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
  // https://docs.microsoft.com/previous-versions/office/office-365-api/api/version-2.0/mail-rest-operations#get-messages.
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
