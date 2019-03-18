---
ms.date: 02/08/2019
description: Lide com erros nas funções personalizadas do Excel.
title: Tratamento de erros para funções personalizadas no Excel (versão prévia)
localization_priority: Priority
ms.openlocfilehash: 170da03331663d6779bed7bf0bf5a9b75b908b3f
ms.sourcegitcommit: 8fb60c3a31faedaea8b51b46238eb80c590a2491
ms.translationtype: HT
ms.contentlocale: pt-BR
ms.lasthandoff: 03/14/2019
ms.locfileid: "30632692"
---
# <a name="error-handling-within-custom-functions"></a><span data-ttu-id="69afb-103">Tratamento de erros nas funções personalizadas</span><span class="sxs-lookup"><span data-stu-id="69afb-103">Error handling within custom functions</span></span>

<span data-ttu-id="69afb-104">Quando você cria um suplemento que define funções personalizadas certifique-se de incluir a lógica de tratamento de erro para lidar com os erros de tempo de execução.</span><span class="sxs-lookup"><span data-stu-id="69afb-104">When you build an add-in that defines custom functions, be sure to include error handling logic to account for runtime errors.</span></span> <span data-ttu-id="69afb-105">O tratamento de erro para funções personalizadas equivale  ao [tratamento de erro para API JavaScript do Excel em](excel-add-ins-error-handling.md).</span><span class="sxs-lookup"><span data-stu-id="69afb-105">Error handling for custom functions is the same as [error handling for the Excel JavaScript API at large](excel-add-ins-error-handling.md).</span></span>

[!include[Excel custom functions note](../includes/excel-custom-functions-note.md)]

<span data-ttu-id="69afb-106">No seguinte exemplo de código `.catch` tratará os erros que ocorreram anteriormente no código.</span><span class="sxs-lookup"><span data-stu-id="69afb-106">In the following code sample, `.catch` will handle any errors that occur previously in the code.</span></span>

```js
function getComment(commentID) {
  let url = "https://www.contoso.com/comments/" + x;

  return fetch(url)
    .then(function (data) {
      return data.json();
    })
    .then((json) => {
      return json.body;
    })
    .catch(function (error) {
      throw error;
    })
}
```

## <a name="see-also"></a><span data-ttu-id="69afb-107">Confira também</span><span class="sxs-lookup"><span data-stu-id="69afb-107">See also</span></span>

* [<span data-ttu-id="69afb-108">Tutorial de funções personalizadas do Excel</span><span class="sxs-lookup"><span data-stu-id="69afb-108">Excel custom functions tutorial</span></span>](../tutorials/excel-tutorial-create-custom-functions.md)
* [<span data-ttu-id="69afb-109">Metadados de funções personalizadas</span><span class="sxs-lookup"><span data-stu-id="69afb-109">Custom functions metadata</span></span>](custom-functions-json.md)
* [<span data-ttu-id="69afb-110">Tempo de execução de funções personalizadas do Excel</span><span class="sxs-lookup"><span data-stu-id="69afb-110">Runtime for Excel custom functions</span></span>](custom-functions-runtime.md)
* <span data-ttu-id="69afb-111">[Práticas recomendadas de funções personalizadas](custom-functions-best-practices.md).</span><span class="sxs-lookup"><span data-stu-id="69afb-111">[Custom functions best practices](custom-functions-best-practices.md)</span></span>
* [<span data-ttu-id="69afb-112">Log de alteração de funções personalizadas</span><span class="sxs-lookup"><span data-stu-id="69afb-112">Custom functions changelog</span></span>](custom-functions-changelog.md)
