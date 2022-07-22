---
ms.date: 06/15/2022
description: Entenda as funções personalizadas do Excel que não usam um runtime compartilhado e seu runtime específico do JavaScript.
title: Tempo de execução somente de JavaScript para funções personalizadas
ms.localizationpriority: medium
ms.openlocfilehash: 0d3298e95ab39f976c3fbfd5c0cc4ecdd1369721
ms.sourcegitcommit: b6a3815a1ad17f3522ca35247a3fd5d7105e174e
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 07/22/2022
ms.locfileid: "66958405"
---
# <a name="javascript-only-runtime-for-custom-functions"></a>Tempo de execução somente de JavaScript para funções personalizadas

As funções personalizadas que não usam um runtime compartilhado usam um runtime somente JavaScript projetado para otimizar o desempenho dos cálculos.

[!include[Excel custom functions note](../includes/excel-custom-functions-note.md)]

[!include[Shared runtime note](../includes/shared-runtime-note.md)]

Esse runtime do JavaScript fornece acesso a APIs `OfficeRuntime` no namespace que pode ser usado por funções personalizadas e o painel de tarefas (que é executado em um runtime diferente) para armazenar dados.

## <a name="request-external-data"></a>Solicitar dados externos

É possível solicitar dados externos em uma função personalizada por meio de uma API, como a API [Fetch](https://developer.mozilla.org/docs/Web/API/Fetch_API), ou por meio de um objeto [XmlHttpRequest (XHR)](https://developer.mozilla.org/docs/Web/API/XMLHttpRequest), uma API Web padrão que envia solicitações HTTP para interagir com os servidores.

Lembre-se de que as funções personalizadas devem usar medidas de segurança adicionais ao fazer XmlHttpRequests, exigindo a Mesma Política de [Origem e CORS simples](https://www.w3.org/TR/cors/).[](https://developer.mozilla.org/docs/Web/Security/Same-origin_policy)

Uma implementação de CORS simples não pode usar cookies e dá suporte apenas a métodos simples (GET, HEAD, POST). A CORS simples aceita cabeçalhos simples com nomes de campos `Accept`, `Accept-Language`, `Content-Language`. Você também pode usar um `Content-Type` cabeçalho em CORS simples, desde que o tipo de conteúdo seja `application/x-www-form-urlencoded`, `text/plain`ou `multipart/form-data`.

## <a name="store-and-access-data"></a>Armazenar e acessar dados

Em uma função personalizada que não usa um runtime compartilhado, você pode armazenar e acessar dados usando o objeto [OfficeRuntime.storage](/javascript/api/office-runtime/officeruntime.storage) . `Storage` O objeto é um sistema de armazenamento de chave-valor persistente, não criptografado que fornece uma alternativa ao [localStorage](https://developer.mozilla.org/docs/Web/API/Window/localStorage), que não pode ser usado por funções personalizadas que usam o runtime somente JavaScript. O `Storage` objeto oferece 10 MB de dados por domínio. Os domínios podem ser compartilhados por mais de um suplemento.

O `Storage` objeto é uma solução de armazenamento compartilhado, o que significa que várias partes de um suplemento são capazes de acessar os mesmos dados. Por exemplo, os tokens para autenticação de usuário podem ser armazenados no objeto porque podem ser acessados `Storage` por uma função personalizada (usando o runtime somente JavaScript) e um painel de tarefas (usando um runtime completo do Webview). Da mesma forma, se dois suplementos compartilharem o mesmo domínio (por exemplo, `www.contoso.com/addin1`, ), `www.contoso.com/addin2`eles também têm permissão para compartilhar informações por meio do `Storage` objeto. Observe que os suplementos que têm subdomínios diferentes terão instâncias diferentes `Storage` (por exemplo, `subdomain.contoso.com/addin1`, `differentsubdomain.contoso.com/addin2`).

Como o `Storage` objeto pode ser um local compartilhado, é importante perceber que é possível substituir pares chave-valor.

Os métodos a seguir estão disponíveis no `Storage` objeto.

- `getItem`
- `getItems`
- `setItem`
- `setItems`
- `removeItem`
- `removeItems`
- `getKeys`

> [!NOTE]
> Não há nenhum método para limpar todas as informações (como `clear`). Em vez disso, use `removeItems` para remover várias entradas de uma só vez.

### <a name="officeruntimestorage-example"></a>Exemplo de OfficeRuntime.storage

O exemplo de código a seguir chama o `OfficeRuntime.storage.setItem` método para definir uma chave e um valor em `storage`.

```js
function StoreValue(key, value) {

  return OfficeRuntime.storage.setItem(key, value).then(function (result) {
      return "Success: Item with key '" + key + "' saved to storage.";
  }, function (error) {
      return "Error: Unable to save item with key '" + key + "' to storage. " + error;
  });
}
```

## <a name="next-steps"></a>Próximas etapas

Saiba como [depurar funções personalizadas](custom-functions-debugging.md).

## <a name="see-also"></a>Confira também

- [Autenticação para funções personalizadas sem um tempo de execução compartilhado](custom-functions-authentication.md)
- [Criar funções personalizadas no Excel](custom-functions-overview.md)
- [Tutorial de funções personalizadas](../tutorials/excel-tutorial-create-custom-functions.md)
