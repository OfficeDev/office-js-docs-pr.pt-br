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
# <a name="display-a-dialog-box-in-custom-functions"></a>Exibir uma caixa de diálogo em funções personalizadas

Se precisar que sua função personalizada interaja com o usuário, você pode criar uma caixa de diálogo usando o objeto `OfficeRuntime.Dialog`. Um cenário comum para usar a caixa de diálogo é autenticar um usuário para que a função personalizada possa acessar um serviço web. Para saber mais sobre autenticação de funções personalizadas, confira [Autenticação de funções personalizados](./custom-functions-authentication.md).

Observação: O objeto `OfficeRuntime.Dialog` faz parte do tempo de execução de funções personalizadas. Ele não pode ser usado a partir do contexto do painel de tarefas. Para criar uma caixa de diálogo de um painel de tarefas, confira [Diálogo API](/office/dev/add-ins/develop/dialog-api-in-office-add-ins).

## <a name="dialog-api-example"></a>Exemplo de API da caixa de diálogo

No exemplo de código a seguir, a função `getTokenViaDialog` usa a função `displayWebDialog` da API da caixa de diálogo para exibir uma caixa de diálogo.

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

## <a name="see-also"></a>Confira também

* [Metadados de funções personalizadas](custom-functions-json.md)
* [Tempo de execução de funções personalizadas do Excel](custom-functions-runtime.md)
* [Práticas recomendadas de funções personalizadas](custom-functions-best-practices.md).
* [Log de alteração de funções personalizadas](custom-functions-changelog.md)
* [Tutorial de funções personalizadas do Excel](../tutorials/excel-tutorial-create-custom-functions.md)
