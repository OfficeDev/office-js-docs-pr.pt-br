---
ms.date: 03/21/2019
description: Crie caixas de diálogo por meio de funções personalizadas no Excel usando JavaScript.
title: Caixas de diálogo de funções personalizados (prévia)
localization_priority: Priority
ms.openlocfilehash: 0f596825a7a32525a68ef45656f1390196146706
ms.sourcegitcommit: a2950492a2337de3180b713f5693fe82dbdd6a17
ms.translationtype: HT
ms.contentlocale: pt-BR
ms.lasthandoff: 03/27/2019
ms.locfileid: "30926640"
---
# <a name="display-a-dialog-box-in-custom-functions"></a><span data-ttu-id="85d5b-103">Exibir uma caixa de diálogo em funções personalizadas</span><span class="sxs-lookup"><span data-stu-id="85d5b-103">Display a dialog box in custom functions</span></span>

<span data-ttu-id="85d5b-104">Se precisar que sua função personalizada interaja com o usuário, você pode criar uma caixa de diálogo usando o objeto `OfficeRuntime.Dialog`.</span><span class="sxs-lookup"><span data-stu-id="85d5b-104">If your custom function needs to interact with the user, you can create a dialog box using the `OfficeRuntime.Dialog` object.</span></span> <span data-ttu-id="85d5b-105">Um cenário comum para usar a caixa de diálogo é autenticar um usuário para que a função personalizada possa acessar um serviço web.</span><span class="sxs-lookup"><span data-stu-id="85d5b-105">A common scenario for using the dialog box is to authenticate a user so that your custom function can access a web service.</span></span> <span data-ttu-id="85d5b-106">Para saber mais sobre autenticação de funções personalizadas, confira [Autenticação de funções personalizados](./custom-functions-authentication.md).</span><span class="sxs-lookup"><span data-stu-id="85d5b-106">For more details about authentication with custom functions, see [Custom functions authentication](./custom-functions-authentication.md).</span></span>

<span data-ttu-id="85d5b-107">Observação: O objeto `OfficeRuntime.Dialog` faz parte do tempo de execução de funções personalizadas.</span><span class="sxs-lookup"><span data-stu-id="85d5b-107">Note: The `OfficeRuntime.Dialog` object is part of the custom functions runtime.</span></span> <span data-ttu-id="85d5b-108">Ele não pode ser usado a partir do contexto do painel de tarefas.</span><span class="sxs-lookup"><span data-stu-id="85d5b-108">It cannot be used from the context of a task pane.</span></span> <span data-ttu-id="85d5b-109">Para criar uma caixa de diálogo de um painel de tarefas, confira [Diálogo API](/office/dev/add-ins/develop/dialog-api-in-office-add-ins).</span><span class="sxs-lookup"><span data-stu-id="85d5b-109">To create a dialog from a task pane, see [Dialog API](/office/dev/add-ins/develop/dialog-api-in-office-add-ins).</span></span>

## <a name="dialog-api-example"></a><span data-ttu-id="85d5b-110">Exemplo de API da caixa de diálogo</span><span class="sxs-lookup"><span data-stu-id="85d5b-110">Dialog API example</span></span>

<span data-ttu-id="85d5b-111">No exemplo de código a seguir, a função `getTokenViaDialog` usa a função `displayWebDialog` da API da caixa de diálogo para exibir uma caixa de diálogo.</span><span class="sxs-lookup"><span data-stu-id="85d5b-111">In the following code sample, the function `getTokenViaDialog` uses the Dialog API’s `displayWebDialog` function to display a dialog box.</span></span>

```js
// Get auth token before calling my service, a hypothetical API that will deliver a stock price based on stock ticker string, such as "MSFT"

function getStock (ticker) {
  return new Promise(function (resolve, reject) {
    // Get a token
    getToken("https://www.contoso.com/auth")
    .then(function (token) {

      // Use token to get stock price
      fetch("https://www.contoso.com/?token=token&ticker= + ticker")
      .then(function (result) {

        // Return stock price to cell
        resolve(result);
      });
    })
    .catch(function (error) {
      reject(error);
    });
  });
  
  //Helper
  function getToken(url) {
    return new Promise(function (resolve,reject) {
      if(_cachedToken) {
        resolve(_cachedToken);
      } else {
        getTokenViaDialog(url)
        .then(function (result) {
          resolve(result);
        })
        .catch(function (result) {
          reject(result);
        });
      }
    });
  }

  function getTokenViaDialog(url) {
    return new Promise (function (resolve, reject) {
      if (_dialogOpen) {
        // Can only have one dialog open at once, wait for previous dialog's token
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
}
```

## <a name="see-also"></a><span data-ttu-id="85d5b-112">Confira também</span><span class="sxs-lookup"><span data-stu-id="85d5b-112">See also</span></span>

* [<span data-ttu-id="85d5b-113">Metadados de funções personalizadas</span><span class="sxs-lookup"><span data-stu-id="85d5b-113">Custom functions metadata</span></span>](custom-functions-json.md)
* [<span data-ttu-id="85d5b-114">Tempo de execução de funções personalizadas do Excel</span><span class="sxs-lookup"><span data-stu-id="85d5b-114">Runtime for Excel custom functions</span></span>](custom-functions-runtime.md)
* <span data-ttu-id="85d5b-115">[Práticas recomendadas de funções personalizadas](custom-functions-best-practices.md).</span><span class="sxs-lookup"><span data-stu-id="85d5b-115">[Custom functions best practices](custom-functions-best-practices.md)</span></span>
* [<span data-ttu-id="85d5b-116">Log de alteração de funções personalizadas</span><span class="sxs-lookup"><span data-stu-id="85d5b-116">Custom functions changelog</span></span>](custom-functions-changelog.md)
* [<span data-ttu-id="85d5b-117">Tutorial de funções personalizadas do Excel</span><span class="sxs-lookup"><span data-stu-id="85d5b-117">Excel custom functions tutorial</span></span>](../tutorials/excel-tutorial-create-custom-functions.md)
