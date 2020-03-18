---
ms.date: 06/18/2019
description: Crie uma caixa de diálogo por meio de funções personalizadas no Excel usando JavaScript.
title: Exibir uma caixa de diálogo a partir de um função personalizada
localization_priority: Normal
ms.openlocfilehash: a2ef005f4c1519228f114dbd671d689807e5914c
ms.sourcegitcommit: fa4e81fcf41b1c39d5516edf078f3ffdbd4a3997
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 03/17/2020
ms.locfileid: "42718739"
---
# <a name="display-a-dialog-box-from-a-custom-function"></a><span data-ttu-id="64107-103">Exibir uma caixa de diálogo a partir de um função personalizada</span><span class="sxs-lookup"><span data-stu-id="64107-103">Display a dialog box from a custom function</span></span>

<span data-ttu-id="64107-104">Se precisar que sua função personalizada interaja com o usuário, você pode criar uma caixa de diálogo usando o [`Office.Dialog`objeto](/javascript/api/office-runtime/officeruntime.dialog).</span><span class="sxs-lookup"><span data-stu-id="64107-104">If your custom function needs to interact with the user, you can create a dialog box using the [`Office.Dialog` object](/javascript/api/office-runtime/officeruntime.dialog).</span></span> <span data-ttu-id="64107-105">Um cenário comum para usar a caixa de diálogo é autenticar um usuário para que a função personalizada possa acessar um serviço web.</span><span class="sxs-lookup"><span data-stu-id="64107-105">A common scenario for using the dialog box is to authenticate a user so that your custom function can access a web service.</span></span> <span data-ttu-id="64107-106">Para saber mais sobre autenticação de funções personalizadas, confira [Autenticação de funções personalizados](./custom-functions-authentication.md).</span><span class="sxs-lookup"><span data-stu-id="64107-106">For more details about authentication with custom functions, see [Custom functions authentication](./custom-functions-authentication.md).</span></span>

[!include[Excel custom functions note](../includes/excel-custom-functions-note.md)]

>[!NOTE]
> <span data-ttu-id="64107-107">O objeto `Office.Dialog` faz parte do tempo de execução de funções personalizadas.</span><span class="sxs-lookup"><span data-stu-id="64107-107">The `Office.Dialog` object is part of the custom functions runtime.</span></span> <span data-ttu-id="64107-108">Painéis de tarefas não usam o objeto `Dialog`.</span><span class="sxs-lookup"><span data-stu-id="64107-108">Task panes don't use the `Dialog` object.</span></span> <span data-ttu-id="64107-109">Para criar uma caixa de diálogo a partir de um painel de tarefas, confira [API de Caixa de Diálogo](../develop/dialog-api-in-office-add-ins.md).</span><span class="sxs-lookup"><span data-stu-id="64107-109">To create a dialog box from a task pane, see [Dialog API](../develop/dialog-api-in-office-add-ins.md).</span></span>

## <a name="dialog-box-api-example"></a><span data-ttu-id="64107-110">exemplo de API da caixa de diálogo</span><span class="sxs-lookup"><span data-stu-id="64107-110">dialog box API example</span></span>

<span data-ttu-id="64107-111">No exemplo de código a seguir, a `getTokenViaDialog` função usa `Dialog` a função `displayWebDialogOptions` da API para exibir uma caixa de diálogo.</span><span class="sxs-lookup"><span data-stu-id="64107-111">In the following code sample, the function `getTokenViaDialog` uses the `Dialog` API's `displayWebDialogOptions` function to display a dialog box.</span></span>

```js
/**
 * Function retrieves a cached token or opens a dialog box if there is no saved token. Note that this is not a sufficient example of authentication but is intended to show the capabilities of the Dialog object.
 * @param {string} url URL for a stored token.
 */
function getTokenViaDialog(url) {
  return new Promise (function (resolve, reject) {
    if (_dialogOpen) {
      // Can only have one dialog box open at once, wait for previous dialog box's token
      let timeout = 5;
      let count = 0;
      var intervalId = setInterval(function () {
        count++;
        if(_cachedToken) {
          resolve(_cachedToken);
          clearInterval(intervalId);
        }
        if(count >= timeout) {
          reject("Timeout while waiting for token");
          clearInterval(intervalId);
        }
      }, 1000);
    } else {
      _dialogOpen = true;
      OfficeRuntime.displayWebDialog(url, {
        height: '50%',
        width: '50%',
        onMessage: function (message, dialog) {
          _cachedToken = message;
          resolve(message);
          dialog.close();
          return;
        },
        onRuntimeError: function(error, dialog) {
          reject(error);
        },
      }).catch(function (e) {
        reject(e);
      });
    }
  });
}
```

## <a name="next-steps"></a><span data-ttu-id="64107-112">Próximas etapas</span><span class="sxs-lookup"><span data-stu-id="64107-112">Next steps</span></span>
<span data-ttu-id="64107-113">Saiba como [tornar as suas funções personalizadas compatíveis com as funções definidas pelo usuário de XLL](make-custom-functions-compatible-with-xll-udf.md).</span><span class="sxs-lookup"><span data-stu-id="64107-113">Learn how to [make your custom functions compatible with XLL user-defined functions](make-custom-functions-compatible-with-xll-udf.md).</span></span>

## <a name="see-also"></a><span data-ttu-id="64107-114">Confira também</span><span class="sxs-lookup"><span data-stu-id="64107-114">See also</span></span>

* [<span data-ttu-id="64107-115">Autenticação de funções personalizadas</span><span class="sxs-lookup"><span data-stu-id="64107-115">Custom functions authentication</span></span>](custom-functions-authentication.md)
* [<span data-ttu-id="64107-116">Receber e tratar dados com funções personalizadas</span><span class="sxs-lookup"><span data-stu-id="64107-116">Receive and handle data with custom functions</span></span>](custom-functions-web-reqs.md)
* [<span data-ttu-id="64107-117">Criar funções personalizadas no Excel</span><span class="sxs-lookup"><span data-stu-id="64107-117">Create custom functions in Excel</span></span>](custom-functions-overview.md)
