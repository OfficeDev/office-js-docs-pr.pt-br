---
ms.date: 02/08/2019
description: Lide com erros nas funções personalizadas do Excel.
title: Tratamento de erros para funções personalizadas no Excel (versão prévia)
localization_priority: Priority
ms.openlocfilehash: 6c1c7f780aea125977510e4eb0e320933cd6ed9c
ms.sourcegitcommit: 9e7b4daa8d76c710b9d9dd4ae2e3c45e8fe07127
ms.translationtype: HT
ms.contentlocale: pt-BR
ms.lasthandoff: 04/24/2019
ms.locfileid: "32448319"
---
# <a name="error-handling-within-custom-functions"></a><span data-ttu-id="34ebc-103">Tratamento de erros nas funções personalizadas</span><span class="sxs-lookup"><span data-stu-id="34ebc-103">Error handling within custom functions</span></span>

<span data-ttu-id="34ebc-104">Quando você cria um suplemento que define funções personalizadas certifique-se de incluir a lógica de tratamento de erro para lidar com os erros de tempo de execução.</span><span class="sxs-lookup"><span data-stu-id="34ebc-104">When you build an add-in that defines custom functions, be sure to include error handling logic to account for runtime errors.</span></span> <span data-ttu-id="34ebc-105">O tratamento de erro para funções personalizadas equivale  ao [tratamento de erro para API JavaScript do Excel em](excel-add-ins-error-handling.md).</span><span class="sxs-lookup"><span data-stu-id="34ebc-105">Error handling for custom functions is the same as [error handling for the Excel JavaScript API at large](excel-add-ins-error-handling.md).</span></span>

[!include[Excel custom functions note](../includes/excel-custom-functions-note.md)]

<span data-ttu-id="34ebc-106">No seguinte exemplo de código `.catch` tratará os erros que ocorreram anteriormente no código.</span><span class="sxs-lookup"><span data-stu-id="34ebc-106">In the following code sample, `.catch` will handle any errors that occur previously in the code.</span></span>

```js
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

## <a name="see-also"></a><span data-ttu-id="34ebc-107">Confira também</span><span class="sxs-lookup"><span data-stu-id="34ebc-107">See also</span></span>

* [<span data-ttu-id="34ebc-108">Tutorial de funções personalizadas do Excel</span><span class="sxs-lookup"><span data-stu-id="34ebc-108">Excel custom functions tutorial</span></span>](../tutorials/excel-tutorial-create-custom-functions.md)
* [<span data-ttu-id="34ebc-109">Metadados de funções personalizadas</span><span class="sxs-lookup"><span data-stu-id="34ebc-109">Custom functions metadata</span></span>](custom-functions-json.md)
* [<span data-ttu-id="34ebc-110">Tempo de execução de funções personalizadas do Excel</span><span class="sxs-lookup"><span data-stu-id="34ebc-110">Runtime for Excel custom functions</span></span>](custom-functions-runtime.md)
* <span data-ttu-id="34ebc-111">[Práticas recomendadas de funções personalizadas](custom-functions-best-practices.md).</span><span class="sxs-lookup"><span data-stu-id="34ebc-111">[Custom functions best practices](custom-functions-best-practices.md)</span></span>
* [<span data-ttu-id="34ebc-112">Log de alteração de funções personalizadas</span><span class="sxs-lookup"><span data-stu-id="34ebc-112">Custom functions changelog</span></span>](custom-functions-changelog.md)
