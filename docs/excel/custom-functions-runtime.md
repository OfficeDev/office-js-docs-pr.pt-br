---
ms.date: 05/17/2020
description: Entenda as funções personalizadas do Excel que não usam um painel de tarefas e seu tempo de execução JavaScript específico.
title: Tempo de execução para funções personalizadas do Excel sem interface do usuário
localization_priority: Normal
ms.openlocfilehash: 31044d4569d230e252c05a39785fc7d47b802e37
ms.sourcegitcommit: f62d9630de69c5c070e3d4048205f5cc654db7e4
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 05/18/2020
ms.locfileid: "44278354"
---
# <a name="runtime-for-ui-less-excel-custom-functions"></a>Tempo de execução para funções personalizadas do Excel sem interface do usuário

As funções personalizadas que não usam um painel de tarefas (funções personalizadas sem interface do usuário) usam um tempo de execução do JavaScript projetado para otimizar o desempenho dos cálculos.

[!include[Excel custom functions note](../includes/excel-custom-functions-note.md)]

[!include[Shared runtime note](../includes/shared-runtime-note.md)]

Este tempo de execução JavaScript fornece acesso a APIs no `OfficeRuntime` namespace que podem ser usadas por funções personalizadas sem interface do usuário e o painel de tarefas para armazenar dados.

## <a name="requesting-external-data"></a>Como solicitar dados externos

Dentro de uma função personalizada sem interface do usuário, você pode solicitar dados externos usando uma API como [Fetch](https://developer.mozilla.org/en-US/docs/Web/API/Fetch_API) ou usando [XMLHttpRequest (XHR)](https://developer.mozilla.org/en-US/docs/Web/API/XMLHttpRequest), uma API Web padrão que emite solicitações HTTP para interagir com os servidores.

Esteja ciente de que as funções sem interface do usuário devem usar medidas de segurança adicionais ao fazer XMLHttpRequests, exigindo a [mesma política de origem](https://developer.mozilla.org/en-US/docs/Web/Security/Same-origin_policy) e [CORS](https://www.w3.org/TR/cors/)simples.

Uma implementação CORS simples não pode usar cookies e só oferece suporte a métodos simples (GET, HEAD, POST). A CORS simples aceita cabeçalhos simples com nomes de campos `Accept`, `Accept-Language`, `Content-Language`. Você também pode usar um `Content-Type` cabeçalho no CORS simples, desde que o tipo de conteúdo seja `application/x-www-form-urlencoded` , `text/plain` ou `multipart/form-data` .

## <a name="storing-and-accessing-data"></a>Como armazenar e acessar os dados

Dentro de uma função personalizada sem interface do usuário, você pode armazenar e acessar dados usando o `OfficeRuntime.storage` objeto. `Storage`é um sistema de armazenamento de valor chave persistente, não criptografado que fornece uma alternativa para o [localStorage](https://developer.mozilla.org/en-US/docs/Web/API/Window/localStorage), que não pode ser usado por funções personalizadas sem interface do usuário. `Storage`o oferece 10 MB de dados por domínio. Os domínios podem ser compartilhados por mais de um suplemento.

`Storage` é uma solução de armazenamento compartilhado, o que significa que várias partes de um suplemento podem acessar os mesmos dados. Por exemplo, os tokens para autenticação de usuário podem ser armazenados em `storage` porque podem ser acessados por uma função personalizada sem interface e elementos de interface do usuário de suplemento, como um painel de tarefas. Da mesma forma, se dois suplementos compartilham o mesmo domínio (por exemplo, `www.contoso.com/addin1` `www.contoso.com/addin2` ), eles também podem compartilhar informações de frente e para trás `storage` . Observe que os suplementos que possuem subdomínios diferentes terão instâncias diferentes `storage` (por exemplo, `subdomain.contoso.com/addin1` , `differentsubdomain.contoso.com/addin2` ).

Como `storage` pode ser um local compartilhado, é importante notar que é possível substituir os pares chave-valor.

Os métodos a seguir estão disponíveis no objeto `storage`:

 - `getItem`
 - `getItems`
 - `setItem`
 - `setItems`
 - `removeItem`
 - `removeItems`
 - `getKeys`

.[!NOTE]
> Não há nenhum método para limpar todas as informações (como `clear` ). Em vez disso, use `removeItems` para remover várias entradas de uma só vez.

### <a name="officeruntimestorage-example"></a>Exemplo de OfficeRuntime. Storage

O exemplo de código a seguir chama a `OfficeRuntime.storage.setItem` função para definir uma chave e um valor para `storage` .

```js
function StoreValue(key, value) {

  return OfficeRuntime.storage.setItem(key, value).then(function (result) {
      return "Success: Item with key '" + key + "' saved to storage.";
  }, function (error) {
      return "Error: Unable to save item with key '" + key + "' to storage. " + error;
  });
}
```

## <a name="additional-considerations"></a>Considerações adicionais

Se o suplemento usar apenas funções personalizadas sem interface do usuário, observe que não é possível acessar o modelo de objeto de documento (DOM) com funções personalizadas sem interface do usuário ou usar bibliotecas como jQuery que dependem do DOM.

## <a name="next-steps"></a>Próximas etapas
Saiba como [depurar funções personalizadas sem interface do usuário](custom-functions-debugging.md).

## <a name="see-also"></a>Confira também

* [Autenticar funções personalizadas sem interface do usuário](custom-functions-authentication.md)
* [Criar funções personalizadas no Excel](custom-functions-overview.md)
* [Tutorial de funções personalizadas](../tutorials/excel-tutorial-create-custom-functions.md)
