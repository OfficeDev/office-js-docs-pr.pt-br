---
ms.date: 06/18/2019
description: Crie uma caixa de diálogo por meio de funções personalizadas no Excel usando JavaScript.
title: Exibir uma caixa de diálogo a partir de um função personalizada
localization_priority: Priority
ms.openlocfilehash: 67a61bde409d45b2c96118de95f0839e7a73ddfe
ms.sourcegitcommit: 654ac1a0c477413662b48cffc0faee5cb65fc25f
ms.translationtype: HT
ms.contentlocale: pt-BR
ms.lasthandoff: 08/09/2019
ms.locfileid: "36268149"
---
# <a name="display-a-dialog-box-from-a-custom-function"></a>Exibir uma caixa de diálogo a partir de um função personalizada

Se precisar que sua função personalizada interaja com o usuário, você pode criar uma caixa de diálogo usando o [`Office.Dialog`objeto](/javascript/api/office-runtime/officeruntime.dialog). Um cenário comum para usar a caixa de diálogo é autenticar um usuário para que a função personalizada possa acessar um serviço web. Para saber mais sobre autenticação de funções personalizadas, confira [Autenticação de funções personalizados](./custom-functions-authentication.md).

[!include[Excel custom functions note](../includes/excel-custom-functions-note.md)]

>[!NOTE]
> O objeto `Office.Dialog` faz parte do tempo de execução de funções personalizadas. Painéis de tarefas não usam o objeto `Dialog`. Para criar uma caixa de diálogo a partir de um painel de tarefas, confira [API de Caixa de Diálogo](/office/dev/add-ins/develop/dialog-api-in-office-add-ins).

## <a name="dialog-box-api-example"></a>exemplo de API da caixa de diálogo

Na amostra de código a seguir, a função `getTokenViaDialog` usa a função da `Dialog`API`displayWebDialogOptions` para exibir uma caixa de diálogo.

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

## <a name="next-steps"></a>Próximas etapas
Saiba como [tornar as suas funções personalizadas compatíveis com as funções definidas pelo usuário de XLL](make-custom-functions-compatible-with-xll-udf.md).

## <a name="see-also"></a>Confira também

* [Autenticação de funções personalizadas](custom-functions-authentication.md)
* [Receber e tratar dados com funções personalizadas](custom-functions-web-reqs.md)
* [Criar funções personalizadas no Excel](custom-functions-overview.md)
