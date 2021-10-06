---
title: Elemento FunctionFile no arquivo de manifesto
description: Especifica o arquivo de código-fonte para operações expostas por um suplemento através de comandos de suplemento que executam uma função JavaScript, em vez de exibir a interface do usuário.
ms.date: 09/29/2021
ms.localizationpriority: medium
ms.openlocfilehash: e8d65e8d8ba94dd63dc82c0519260157b1d22a62
ms.sourcegitcommit: 489befc41e543a4fb3c504fd9b3f61322134c1ef
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 10/06/2021
ms.locfileid: "60138755"
---
# <a name="functionfile-element"></a>Elemento FunctionFile

Especifica o arquivo de código-fonte para operações que um complemento expõe de uma das seguintes maneiras.

* Comandos de complemento que executam uma função JavaScript em vez de exibir a interface do usuário.
* Atalhos de teclado que executam uma função JavaScript.

**Tipo de complemento:** Painel de tarefas, Email

**Válido somente nestes esquemas VersionOverrides:**

- Painel de tarefas 1.0
- Email 1.0
- Email 1.1

Para obter mais informações, consulte [Substituições de versão no manifesto](../../develop/add-in-manifests.md#version-overrides-in-the-manifest).

O `FunctionFile` elemento é um elemento filho de [DesktopFormFactor](desktopformfactor.md) ou [MobileFormFactor](mobileformfactor.md). O atributo do elemento não pode ter mais de 32 caracteres e é definido como o valor do atributo de um elemento no elemento que contém a URL para um arquivo HTML que contém ou carrega todas as funções `resid` `FunctionFile` `id` `Url` `Resources` JavaScript usadas [](control.md)por botões de comando de complemento sem interface do usuário, conforme definido pelo elemento Control .

> [!NOTE]
> Quando o add-in é configurado para usar um tempo de execução [compartilhado,](../../develop/configure-your-add-in-to-use-a-shared-runtime.md)as funções no arquivo de código são executados no mesmo tempo de execução javascript (e compartilham um namespace global comum) que o JavaScript no painel de tarefas do complemento (se for o caso).
>
> O elemento e o arquivo de código associado também têm uma função especial para executar com atalhos de teclado `FunctionFile` [personalizados,](../../design/keyboard-shortcuts.md)que exigem um tempo de execução compartilhado.

A seguir, um exemplo do `FunctionFile` elemento.

```XML
<DesktopFormFactor>
  <FunctionFile resid="residDesktopFuncUrl" />
  <ExtensionPoint xsi:type="PrimaryCommandSurface">
    <!-- information about this extension point -->
  </ExtensionPoint>

  <!-- You can define more than one ExtensionPoint element as needed -->

</DesktopFormFactor>
```

O JavaScript no arquivo HTML indicado pelo elemento deve chamar e definir funções nomeadas que têm `FunctionFile` `Office.initialize` um único parâmetro: `event` . As funções devem usar a API `item.notificationMessages` para indicar o progresso, sucesso ou falha ao usuário. Também deverá chamar `event.completed` quando terminar a execução. O nome das funções é usado no `FunctionName` elemento para botões sem interface do usuário.

A seguir, um exemplo de um arquivo HTML que define uma `trackMessage` função.

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

O código a seguir mostra como implementar a função usada por `FunctionName` .

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
> A chamada para `event.completed` sinais de que você lidou com êxito com o evento. Quando uma função é chamada várias vezes, por exemplo, com vários cliques no mesmo comando de suplemento, todos os eventos são enfileirados automaticamente. O primeiro evento é executado automaticamente, enquanto os outros eventos permanecem na fila. Quando sua função chama , a próxima chamada `event.completed` em fila para essa função é executado. Você deve `event.completed` chamar; caso contrário, sua função não será executado.
