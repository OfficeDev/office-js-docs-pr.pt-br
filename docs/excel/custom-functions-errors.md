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
# <a name="error-handling-within-custom-functions"></a><span data-ttu-id="1ff57-103">Tratamento de erros nas funções personalizadas</span><span class="sxs-lookup"><span data-stu-id="1ff57-103">Error handling within custom functions</span></span>

<span data-ttu-id="1ff57-104">Quando você cria um suplemento que define funções personalizadas certifique-se de incluir a lógica de tratamento de erro para lidar com os erros de tempo de execução.</span><span class="sxs-lookup"><span data-stu-id="1ff57-104">When you build an add-in that defines custom functions, be sure to include error handling logic to account for runtime errors.</span></span> <span data-ttu-id="1ff57-105">O tratamento de erro para funções personalizadas equivale  ao [tratamento de erro para API JavaScript do Excel em](excel-add-ins-error-handling.md).</span><span class="sxs-lookup"><span data-stu-id="1ff57-105">Error handling for custom functions is the same as [error handling for the Excel JavaScript API at large](excel-add-ins-error-handling.md).</span></span>

[!include[Excel custom functions note](../includes/excel-custom-functions-note.md)]

<span data-ttu-id="1ff57-106">No seguinte exemplo de código `.catch` tratará os erros que ocorreram anteriormente no código.</span><span class="sxs-lookup"><span data-stu-id="1ff57-106">In the following code sample, `.catch` will handle any errors that occur previously in the code.</span></span>

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

## <a name="next-steps"></a><span data-ttu-id="1ff57-107">Próximas etapas</span><span class="sxs-lookup"><span data-stu-id="1ff57-107">Next steps</span></span>
<span data-ttu-id="1ff57-108">Saiba como [solucionar problemas com as suas funções personalizadas](custom-functions-troubleshooting.md).</span><span class="sxs-lookup"><span data-stu-id="1ff57-108">Learn how to [troubleshoot problems with your custom functions](custom-functions-troubleshooting.md).</span></span>

## <a name="see-also"></a><span data-ttu-id="1ff57-109">Confira também</span><span class="sxs-lookup"><span data-stu-id="1ff57-109">See also</span></span>

* [<span data-ttu-id="1ff57-110">Depuração de funções personalizadas</span><span class="sxs-lookup"><span data-stu-id="1ff57-110">Custom functions debugging</span></span>](custom-functions-debugging.md)
* [<span data-ttu-id="1ff57-111">Requisitos de funções personalizadas</span><span class="sxs-lookup"><span data-stu-id="1ff57-111">Custom functions requirements</span></span>](custom-functions-requirements.md)
* [<span data-ttu-id="1ff57-112">Criar funções personalizadas no Excel</span><span class="sxs-lookup"><span data-stu-id="1ff57-112">Create custom functions in Excel</span></span>](custom-functions-overview.md)
