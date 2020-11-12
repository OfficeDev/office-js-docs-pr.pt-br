---
title: Elemento FunctionFile no arquivo de manifesto
description: Especifica o arquivo de código-fonte para operações expostas por um suplemento através de comandos de suplemento que executam uma função JavaScript, em vez de exibir a interface do usuário.
ms.date: 11/06/2020
localization_priority: Normal
ms.openlocfilehash: 4c47c3e4b824f2b93aaea17cef88e01f748d6f95
ms.sourcegitcommit: ca66ff7462bfdf4ed7ae04f43d1388c24de63bf9
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 11/11/2020
ms.locfileid: "48996442"
---
# <a name="functionfile-element"></a>Elemento FunctionFile

Especifica o arquivo de código-fonte para as operações que um suplemento expõe de uma das seguintes maneiras:

* Comandos de suplemento que executam uma função JavaScript, em vez de exibir a interface do usuário.
* Atalhos de teclado que executam uma função JavaScript.

O `FunctionFile` elemento é um elemento filho de [DesktopFormFactor](desktopformfactor.md) ou [MobileFormFactor](mobileformfactor.md). O `resid` atributo do `FunctionFile` elemento é definido como o valor do `id` atributo de um `Url` elemento no `Resources` elemento que contém a URL para um arquivo HTML que contém ou carrega todas as funções JavaScript usadas por botões de comando do suplemento sem interface do usuário, conforme definido pelo [elemento Control](control.md).

Veja a seguir um exemplo do `FunctionFile` elemento.

```XML
<DesktopFormFactor>
  <FunctionFile resid="residDesktopFuncUrl" />
  <ExtensionPoint xsi:type="PrimaryCommandSurface">
    <!-- information about this extension point -->
  </ExtensionPoint>

  <!-- You can define more than one ExtensionPoint element as needed -->

</DesktopFormFactor>
```

O JavaScript no arquivo HTML indicado pelo `FunctionFile` elemento deve chamar `Office.initialize` e definir funções nomeadas que usam um único parâmetro: `event` . As funções devem usar a API `item.notificationMessages` para indicar o progresso, sucesso ou falha ao usuário. Também deverá chamar `event.completed` quando terminar a execução. O nome das funções são usados no `FunctionName` elemento para botões sem interface do usuário.

Veja a seguir um exemplo de um arquivo HTML que define uma `trackMessage` função.

```js
Office.initialize = function () {
    doAuth();
}

function trackMessage (event) {
    var buttonId = event.source.id;    
    var itemId = Office.context.mailbox.item.id;
    // save this message
    event.completed();
}
```

O código a seguir mostra como implementar a função usada pelo `FunctionName` .

```js
// The initialize function must be run each time a new page is loaded.
(function () {
    Office.initialize = function (reason) {
        // If you need to initialize something you can do so here.
    };
})();

// Your function must be in the global namespace.
function writeText(event) {

    // Implement your custom code here. The following code is a simple example.

    Office.context.document.setSelectedDataAsync("ExecuteFunction works. Button ID=" + event.source.id,
        function (asyncResult) {
            var error = asyncResult.error;
            if (asyncResult.status === "failed") {
                // Show error message.
            }
            else {
                // Show success message.
            }
        });
    // Calling event.completed is required. event.completed lets the platform know that processing has completed.
    event.completed();
}
```

> [!IMPORTANT]
> A chamada para `event.completed` sinaliza que você tratou com êxito o evento. Quando uma função é chamada várias vezes, por exemplo, com vários cliques no mesmo comando de suplemento, todos os eventos são enfileirados automaticamente. O primeiro evento é executado automaticamente, enquanto os outros eventos permanecem na fila. Quando sua função chama `event.completed` , a próxima chamada em fila para essa função é executada. Você deve chamar `event.completed` ; caso contrário, sua função não será executada.
