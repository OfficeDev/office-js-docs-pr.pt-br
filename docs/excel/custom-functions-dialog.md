---
ms.date: 06/18/2019
description: Crie uma caixa de diálogo por meio de funções personalizadas no Excel usando JavaScript.
title: Exibir uma caixa de diálogo a partir de um função personalizada
localization_priority: Normal
ms.openlocfilehash: 8db5034cf9079ac5cd05654614087882ed1a8d52
ms.sourcegitcommit: d15bca2c12732f8599be2ec4b2adc7c254552f52
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 02/12/2020
ms.locfileid: "41950765"
---
# <a name="display-a-dialog-box-from-a-custom-function"></a><span data-ttu-id="62e4a-103">Exibir uma caixa de diálogo a partir de um função personalizada</span><span class="sxs-lookup"><span data-stu-id="62e4a-103">Display a dialog box from a custom function</span></span>

<span data-ttu-id="62e4a-104">Se precisar que sua função personalizada interaja com o usuário, você pode criar uma caixa de diálogo usando o [`Office.Dialog`objeto](/javascript/api/office-runtime/officeruntime.dialog).</span><span class="sxs-lookup"><span data-stu-id="62e4a-104">If your custom function needs to interact with the user, you can create a dialog box using the [`Office.Dialog` object](/javascript/api/office-runtime/officeruntime.dialog).</span></span> <span data-ttu-id="62e4a-105">Um cenário comum para usar a caixa de diálogo é autenticar um usuário para que a função personalizada possa acessar um serviço web.</span><span class="sxs-lookup"><span data-stu-id="62e4a-105">A common scenario for using the dialog box is to authenticate a user so that your custom function can access a web service.</span></span> <span data-ttu-id="62e4a-106">Para saber mais sobre autenticação de funções personalizadas, confira [Autenticação de funções personalizados](./custom-functions-authentication.md).</span><span class="sxs-lookup"><span data-stu-id="62e4a-106">For more details about authentication with custom functions, see [Custom functions authentication](./custom-functions-authentication.md).</span></span>

[!include[Excel custom functions note](../includes/excel-custom-functions-note.md)]

>[!NOTE]
> <span data-ttu-id="62e4a-107">O objeto `Office.Dialog` faz parte do tempo de execução de funções personalizadas.</span><span class="sxs-lookup"><span data-stu-id="62e4a-107">The `Office.Dialog` object is part of the custom functions runtime.</span></span> <span data-ttu-id="62e4a-108">Painéis de tarefas não usam o objeto `Dialog`.</span><span class="sxs-lookup"><span data-stu-id="62e4a-108">Task panes don't use the `Dialog` object.</span></span> <span data-ttu-id="62e4a-109">Para criar uma caixa de diálogo a partir de um painel de tarefas, confira [API de Caixa de Diálogo](/office/dev/add-ins/develop/dialog-api-in-office-add-ins).</span><span class="sxs-lookup"><span data-stu-id="62e4a-109">To create a dialog box from a task pane, see [Dialog API](/office/dev/add-ins/develop/dialog-api-in-office-add-ins).</span></span>

## <a name="dialog-box-api-example"></a><span data-ttu-id="62e4a-110">exemplo de API da caixa de diálogo</span><span class="sxs-lookup"><span data-stu-id="62e4a-110">dialog box API example</span></span>

<span data-ttu-id="62e4a-111">Na amostra de código a seguir, a função `getTokenViaDialog` usa a função da `Dialog`API`displayWebDialogOptions` para exibir uma caixa de diálogo.</span><span class="sxs-lookup"><span data-stu-id="62e4a-111">In the following code sample, the function `getTokenViaDialog` uses the `Dialog` API’s `displayWebDialogOptions` function to display a dialog box.</span></span>

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

## <a name="next-steps"></a><span data-ttu-id="62e4a-112">Próximas etapas</span><span class="sxs-lookup"><span data-stu-id="62e4a-112">Next steps</span></span>
<span data-ttu-id="62e4a-113">Saiba como [tornar as suas funções personalizadas compatíveis com as funções definidas pelo usuário de XLL](make-custom-functions-compatible-with-xll-udf.md).</span><span class="sxs-lookup"><span data-stu-id="62e4a-113">Learn how to [make your custom functions compatible with XLL user-defined functions](make-custom-functions-compatible-with-xll-udf.md).</span></span>

## <a name="see-also"></a><span data-ttu-id="62e4a-114">Confira também</span><span class="sxs-lookup"><span data-stu-id="62e4a-114">See also</span></span>

* [<span data-ttu-id="62e4a-115">Autenticação de funções personalizadas</span><span class="sxs-lookup"><span data-stu-id="62e4a-115">Custom functions authentication</span></span>](custom-functions-authentication.md)
* [<span data-ttu-id="62e4a-116">Receber e tratar dados com funções personalizadas</span><span class="sxs-lookup"><span data-stu-id="62e4a-116">Receive and handle data with custom functions</span></span>](custom-functions-web-reqs.md)
* [<span data-ttu-id="62e4a-117">Criar funções personalizadas no Excel</span><span class="sxs-lookup"><span data-stu-id="62e4a-117">Create custom functions in Excel</span></span>](custom-functions-overview.md)
