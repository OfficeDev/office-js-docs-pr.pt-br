---
ms.date: 07/08/2021
description: Entenda Excel funções personalizadas que não usam um painel de tarefas e seu tempo de execução JavaScript específico.
title: Tempo de execução para funções personalizadas sem Excel de interface do usuário
ms.localizationpriority: medium
ms.openlocfilehash: 491e47674d87d99d0adeda952ee65ffc24dff2bd
ms.sourcegitcommit: 1306faba8694dea203373972b6ff2e852429a119
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 09/12/2021
ms.locfileid: "59148605"
---
# <a name="runtime-for-ui-less-excel-custom-functions"></a>Tempo de execução para funções personalizadas sem Excel de interface do usuário

Funções personalizadas que não usam um painel de tarefas (funções personalizadas sem interface do usuário) usam um tempo de execução JavaScript projetado para otimizar o desempenho dos cálculos.

[!include[Excel custom functions note](../includes/excel-custom-functions-note.md)]

[!include[Shared runtime note](../includes/shared-runtime-note.md)]

Esse tempo de execução javaScript fornece acesso a APIs no namespace que podem ser usadas por funções personalizadas sem interface do usuário e o painel de tarefas para `OfficeRuntime` armazenar dados.

## <a name="request-external-data"></a>Solicitar dados externos

Em uma função personalizada sem interface do usuário, você pode solicitar dados externos usando uma API como [Fetch](https://developer.mozilla.org/docs/Web/API/Fetch_API) ou usando [XmlHttpRequest (XHR),](https://developer.mozilla.org/docs/Web/API/XMLHttpRequest)uma API Web padrão que emite solicitações HTTP para interagir com servidores.

Esteja ciente de que funções sem interface do usuário devem usar medidas de [](https://developer.mozilla.org/docs/Web/Security/Same-origin_policy) segurança adicionais ao criar XmlHttpRequests, exigindo a Política de Mesma Origem e [o CORS simples.](https://www.w3.org/TR/cors/)

Uma implementação de CORS simples não pode usar cookies e só oferece suporte a métodos simples (GET, HEAD, POST). A CORS simples aceita cabeçalhos simples com nomes de campos `Accept`, `Accept-Language`, `Content-Language`. Você também pode usar `Content-Type` um header em CORS simples, desde que o tipo de conteúdo `application/x-www-form-urlencoded` seja , ou `text/plain` `multipart/form-data` .

## <a name="store-and-access-data"></a>Armazenar e acessar dados

Em uma função personalizada sem interface do usuário, você pode armazenar e acessar dados usando o `OfficeRuntime.storage` objeto. `Storage` é um sistema de armazenamento persistente, não criptografado e de valor-chave que fornece uma alternativa ao [localStorage](https://developer.mozilla.org/docs/Web/API/Window/localStorage), que não pode ser usado por funções personalizadas sem interface do usuário. `Storage` oferece 10 MB de dados por domínio. Os domínios podem ser compartilhados por mais de um complemento.

`Storage` é uma solução de armazenamento compartilhado, o que significa que várias partes de um suplemento podem acessar os mesmos dados. Por exemplo, os tokens para autenticação do usuário podem ser armazenados porque podem ser acessados por uma função personalizada sem interface do usuário e elementos de interface do usuário de complemento, como um `storage` painel de tarefas. Da mesma forma, se dois complementos compartilharem o mesmo domínio (por exemplo, , ), eles também poderão compartilhar informações de ida e `www.contoso.com/addin1` `www.contoso.com/addin2` `storage` volta. Observe que os complementos com subdomas diferentes terão instâncias diferentes `storage` (por exemplo, `subdomain.contoso.com/addin1` , `differentsubdomain.contoso.com/addin2` ).

Como `storage` pode ser um local compartilhado, é importante notar que é possível substituir os pares chave-valor.

Os métodos a seguir estão disponíveis no `storage` objeto.

- `getItem`
- `getItems`
- `setItem`
- `setItems`
- `removeItem`
- `removeItems`
- `getKeys`

> [!NOTE]
> Não há nenhum método para limpar todas as informações (como `clear` ). Em vez disso, use `removeItems` para remover várias entradas de uma só vez.

### <a name="officeruntimestorage-example"></a>Exemplo de OfficeRuntime.storage

O exemplo de código a seguir chama `OfficeRuntime.storage.setItem` a função para definir uma chave e um valor em `storage` .

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

Se o seu complemento usa apenas funções personalizadas sem interface do usuário, observe que você não pode acessar o Dom (Modelo de Objeto de Documento) com funções personalizadas sem interface do usuário ou usar bibliotecas como jQuery que dependem do DOM.

## <a name="next-steps"></a>Próximas etapas

Saiba como [depurar funções personalizadas sem interface do usuário.](custom-functions-debugging.md)

## <a name="see-also"></a>Confira também

* [Autenticar funções personalizadas sem interface do usuário](custom-functions-authentication.md)
* [Criar funções personalizadas no Excel](custom-functions-overview.md)
* [Tutorial de funções personalizadas](../tutorials/excel-tutorial-create-custom-functions.md)
