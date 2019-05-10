---
ms.date: 05/03/2019
description: Lide com erros nas funções personalizadas do Excel.
title: Tratamento de erros para funções personalizadas no Excel
localization_priority: Priority
ms.openlocfilehash: 188ece6c77bc2cafad6f22448fb698e0c0370ef8
ms.sourcegitcommit: ff73cc04e5718765fcbe74181505a974db69c3f5
ms.translationtype: HT
ms.contentlocale: pt-BR
ms.lasthandoff: 05/06/2019
ms.locfileid: "33628155"
---
# <a name="error-handling-within-custom-functions"></a>Tratamento de erros nas funções personalizadas

Quando você cria um suplemento que define funções personalizadas certifique-se de incluir a lógica de tratamento de erro para lidar com os erros de tempo de execução. O tratamento de erro para funções personalizadas equivale  ao [tratamento de erro para API JavaScript do Excel em](excel-add-ins-error-handling.md).

[!include[Excel custom functions note](../includes/excel-custom-functions-note.md)]

No seguinte exemplo de código `.catch` tratará os erros que ocorreram anteriormente no código.

```js
/**
 * Gets a comment from the hypothetical contoso.com/comments API.
 * @customfunction
 * @param {number} commentID ID of a comment.
 */
function getComment(commentID) {
  let url = "https://www.contoso.com/comments/" + x;

  return fetch(url)
    .then(function (data) {
      return data.json();
    })
    .then(function (json) {
      return json.body;
    })
    .catch(function (error) {
      throw error;
    })
}
```

## <a name="next-steps"></a>Próximas etapas
Saiba como [solucionar problemas com as suas funções personalizadas](custom-functions-troubleshooting.md).

## <a name="see-also"></a>Confira também

* [Depuração de funções personalizadas](custom-functions-debugging.md)
* [Requisitos de funções personalizadas](custom-functions-requirements.md)
* [Criar funções personalizadas no Excel](custom-functions-overview.md)
