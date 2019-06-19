---
ms.date: 06/17/2019
description: Lide com erros nas funções personalizadas do Excel.
title: Tratamento de erros para funções personalizadas no Excel
localization_priority: Priority
ms.openlocfilehash: 5b94d3fc2570eaa310027ebc156aa78c359a56fa
ms.sourcegitcommit: 4bf5159a3821f4277c07d89e88808c4c3a25ff81
ms.translationtype: HT
ms.contentlocale: pt-BR
ms.lasthandoff: 06/18/2019
ms.locfileid: "35059850"
---
# <a name="error-handling-within-custom-functions"></a><span data-ttu-id="d6235-103">Tratamento de erros nas funções personalizadas</span><span class="sxs-lookup"><span data-stu-id="d6235-103">Error handling within custom functions</span></span>

<span data-ttu-id="d6235-104">Quando você cria um suplemento que define funções personalizadas certifique-se de incluir a lógica de tratamento de erro para lidar com os erros de tempo de execução.</span><span class="sxs-lookup"><span data-stu-id="d6235-104">When you build an add-in that defines custom functions, be sure to include error handling logic to account for runtime errors.</span></span> <span data-ttu-id="d6235-105">O tratamento de erro para funções personalizadas equivale  ao [tratamento de erro para API JavaScript do Excel em](excel-add-ins-error-handling.md).</span><span class="sxs-lookup"><span data-stu-id="d6235-105">Error handling for custom functions is the same as [error handling for the Excel JavaScript API at large](excel-add-ins-error-handling.md).</span></span>

<span data-ttu-id="d6235-106">No seguinte exemplo de código `.catch` tratará os erros que ocorreram anteriormente no código.</span><span class="sxs-lookup"><span data-stu-id="d6235-106">In the following code sample, `.catch` will handle any errors that occur previously in the code.</span></span>

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

## <a name="next-steps"></a><span data-ttu-id="d6235-107">Próximas etapas</span><span class="sxs-lookup"><span data-stu-id="d6235-107">Next steps</span></span>
<span data-ttu-id="d6235-108">Saiba como [solucionar problemas com as suas funções personalizadas](custom-functions-troubleshooting.md).</span><span class="sxs-lookup"><span data-stu-id="d6235-108">Learn how to [troubleshoot problems with your custom functions](custom-functions-troubleshooting.md).</span></span>

## <a name="see-also"></a><span data-ttu-id="d6235-109">Confira também</span><span class="sxs-lookup"><span data-stu-id="d6235-109">See also</span></span>

* [<span data-ttu-id="d6235-110">Depuração de funções personalizadas</span><span class="sxs-lookup"><span data-stu-id="d6235-110">Custom functions debugging</span></span>](custom-functions-debugging.md)
* [<span data-ttu-id="d6235-111">Requisitos de funções personalizadas</span><span class="sxs-lookup"><span data-stu-id="d6235-111">Custom functions requirements</span></span>](custom-functions-requirements.md)
* [<span data-ttu-id="d6235-112">Criar funções personalizadas no Excel</span><span class="sxs-lookup"><span data-stu-id="d6235-112">Create custom functions in Excel</span></span>](custom-functions-overview.md)
