---
ms.date: 05/06/2019
description: Crie uma caixa de diálogo por meio de funções personalizadas no Excel usando JavaScript.
title: Exibir uma caixa de diálogo a partir de um função personalizada
localization_priority: Priority
ms.openlocfilehash: 3d7a657402c319b2394c7331b69314b2e5591890
ms.sourcegitcommit: ff73cc04e5718765fcbe74181505a974db69c3f5
ms.translationtype: HT
ms.contentlocale: pt-BR
ms.lasthandoff: 05/06/2019
ms.locfileid: "33628138"
---
# <a name="display-a-dialog-box-from-a-custom-function"></a><span data-ttu-id="053ce-103">Exibir uma caixa de diálogo a partir de um função personalizada</span><span class="sxs-lookup"><span data-stu-id="053ce-103">Display a dialog box from a custom function</span></span>

<span data-ttu-id="053ce-104">Se precisar que sua função personalizada interaja com o usuário, você pode criar uma caixa de diálogo usando o [`Office.Dialog`objeto](/javascript/api/office-runtime/officeruntime.dialog?view=office-js).</span><span class="sxs-lookup"><span data-stu-id="053ce-104">If your custom function needs to interact with the user, you can create a dialog box using the `Office.Dialog` object.</span></span> <span data-ttu-id="053ce-105">Um cenário comum para usar a caixa de diálogo é autenticar um usuário para que a função personalizada possa acessar um serviço web.</span><span class="sxs-lookup"><span data-stu-id="053ce-105">A common scenario for using the dialog box is to authenticate a user so that your custom function can access a web service.</span></span> <span data-ttu-id="053ce-106">Para saber mais sobre autenticação de funções personalizadas, confira [Autenticação de funções personalizados](./custom-functions-authentication.md).</span><span class="sxs-lookup"><span data-stu-id="053ce-106">For more details about authentication with custom functions, see [Custom functions authentication](./custom-functions-authentication.md).</span></span>

[!include[Excel custom functions note](../includes/excel-custom-functions-note.md)]

>[!NOTE]
> <span data-ttu-id="053ce-107">O objeto `Office.Dialog` faz parte do tempo de execução de funções personalizadas.</span><span class="sxs-lookup"><span data-stu-id="053ce-107">Note: The `Office.Dialog` object is part of the custom functions runtime.</span></span> <span data-ttu-id="053ce-108">Painéis de tarefas não usam o objeto `Dialog`.</span><span class="sxs-lookup"><span data-stu-id="053ce-108">Task panes don't use the `Dialog` object.</span></span> <span data-ttu-id="053ce-109">Para criar uma caixa de diálogo a partir de um painel de tarefas, confira [API de Caixa de Diálogo](/office/dev/add-ins/develop/dialog-api-in-office-add-ins).</span><span class="sxs-lookup"><span data-stu-id="053ce-109">To create a dialog from a task pane, see [Dialog API](/office/dev/add-ins/develop/dialog-api-in-office-add-ins).</span></span>

## <a name="dialog-box-api-example"></a><span data-ttu-id="053ce-110">exemplo de API da caixa de diálogo</span><span class="sxs-lookup"><span data-stu-id="053ce-110">Dialog API example</span></span>

<span data-ttu-id="053ce-111">Na amostra de código a seguir, a função `getTokenViaDialog` usa a função da `Dialog`API`displayWebDialogOptions` para exibir uma caixa de diálogo.</span><span class="sxs-lookup"><span data-stu-id="053ce-111">In the following code sample, the function `getTokenViaDialog` uses the Dialog API’s `Dialog` function to display a dialog box.</span></span>

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
      Office.displayWebDialogOptions(url, {
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

## <a name="next-steps"></a><span data-ttu-id="053ce-112">Próximas etapas</span><span class="sxs-lookup"><span data-stu-id="053ce-112">Next steps</span></span>
<span data-ttu-id="053ce-113">Saiba como [tornar as suas funções personalizadas compatíveis com as funções definidas pelo usuário de XLL](make-custom-functions-compatible-with-xll-udf.md).</span><span class="sxs-lookup"><span data-stu-id="053ce-113">Learn how to [make your custom functions compatible with XLL user-defined functions](make-custom-functions-compatible-with-xll-udf.md).</span></span>

## <a name="see-also"></a><span data-ttu-id="053ce-114">Confira também</span><span class="sxs-lookup"><span data-stu-id="053ce-114">See also</span></span>

* <span data-ttu-id="053ce-115">[Autenticação de funções personalizadas](custom-functions-authentication.md)</span><span class="sxs-lookup"><span data-stu-id="053ce-115">For more information, see [Custom functions authentication](custom-functions-authentication.md).</span></span>
* [<span data-ttu-id="053ce-116">Receber e tratar dados com funções personalizadas</span><span class="sxs-lookup"><span data-stu-id="053ce-116">Receive and handle data with custom functions</span></span>](custom-functions-web-reqs.md)
* [<span data-ttu-id="053ce-117">Criar funções personalizadas no Excel</span><span class="sxs-lookup"><span data-stu-id="053ce-117">Create custom functions in Excel</span></span>](custom-functions-overview.md)
