---
title: Elemento FunctionFile no arquivo de manifesto
description: Especifica o arquivo de código-fonte para operações expostas por um suplemento através de comandos de suplemento que executam uma função JavaScript, em vez de exibir a interface do usuário.
ms.date: 10/09/2018
localization_priority: Normal
ms.openlocfilehash: 376ea82f48360d502ea9be05dc5d6b02f9294add
ms.sourcegitcommit: fa4e81fcf41b1c39d5516edf078f3ffdbd4a3997
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 03/17/2020
ms.locfileid: "42718192"
---
# <a name="functionfile-element"></a><span data-ttu-id="2262f-103">Elemento FunctionFile</span><span class="sxs-lookup"><span data-stu-id="2262f-103">FunctionFile element</span></span>

<span data-ttu-id="2262f-104">Especifica o arquivo de código-fonte para operações expostas por um suplemento através de comandos de suplemento que executam uma função JavaScript, em vez de exibir a interface do usuário.</span><span class="sxs-lookup"><span data-stu-id="2262f-104">Specifies the source code file for operations that an add-in exposes through add-in commands that execute a JavaScript function instead of displaying UI.</span></span> <span data-ttu-id="2262f-105">O `FunctionFile` elemento é um elemento filho de [DesktopFormFactor](desktopformfactor.md) ou [MobileFormFactor](mobileformfactor.md).</span><span class="sxs-lookup"><span data-stu-id="2262f-105">The `FunctionFile` element is a child element of [DesktopFormFactor](desktopformfactor.md) or [MobileFormFactor](mobileformfactor.md).</span></span> <span data-ttu-id="2262f-106">O `resid` atributo `FunctionFile` do elemento é definido como o valor do `id` atributo de um `Url` elemento no `Resources` elemento que contém a URL para um arquivo HTML que contém ou carrega todas as funções JavaScript usadas por botões de comando do suplemento sem interface do usuário, conforme definido pelo [elemento Control](control.md).</span><span class="sxs-lookup"><span data-stu-id="2262f-106">The `resid` attribute of the `FunctionFile` element is set to the value of the `id` attribute of a `Url` element in the `Resources` element that contains the URL to an HTML file that contains or loads all  the JavaScript functions used by UI-less add-in command buttons, as defined by the [Control element](control.md).</span></span>

<span data-ttu-id="2262f-107">Veja a seguir um exemplo do `FunctionFile` elemento.</span><span class="sxs-lookup"><span data-stu-id="2262f-107">The following is an example of the `FunctionFile` element.</span></span>

```XML
<DesktopFormFactor>
  <FunctionFile resid="residDesktopFuncUrl" />
  <ExtensionPoint xsi:type="PrimaryCommandSurface">
    <!-- information about this extension point -->
  </ExtensionPoint>

  <!-- You can define more than one ExtensionPoint element as needed -->

</DesktopFormFactor>
```

<span data-ttu-id="2262f-108">O JavaScript no arquivo HTML indicado pelo `FunctionFile` elemento deve chamar `Office.initialize` e definir funções nomeadas que usam um único parâmetro:. `event`</span><span class="sxs-lookup"><span data-stu-id="2262f-108">The JavaScript in the HTML file indicated by the `FunctionFile` element must call `Office.initialize` and define named functions that take a single parameter: `event`.</span></span> <span data-ttu-id="2262f-109">As funções devem usar a API `item.notificationMessages` para indicar o progresso, sucesso ou falha ao usuário.</span><span class="sxs-lookup"><span data-stu-id="2262f-109">The functions should use the `item.notificationMessages` API to indicate progress, success, or failure to the user.</span></span> <span data-ttu-id="2262f-110">Também deverá chamar `event.completed` quando terminar a execução.</span><span class="sxs-lookup"><span data-stu-id="2262f-110">It should also call `event.completed` when it has finished execution.</span></span> <span data-ttu-id="2262f-111">O nome das funções são usados no `FunctionName` elemento para botões sem interface do usuário.</span><span class="sxs-lookup"><span data-stu-id="2262f-111">The name of the functions are used in the `FunctionName` element for UI-less buttons.</span></span>

<span data-ttu-id="2262f-112">Veja a seguir um exemplo de um arquivo HTML que define `trackMessage` uma função.</span><span class="sxs-lookup"><span data-stu-id="2262f-112">The following is an example of an HTML file defining a `trackMessage` function.</span></span>

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

<span data-ttu-id="2262f-113">O código a seguir mostra como implementar a função usada pelo `FunctionName`.</span><span class="sxs-lookup"><span data-stu-id="2262f-113">The following code shows how to implement the function used by `FunctionName`.</span></span>

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
> <span data-ttu-id="2262f-114">A chamada para `event.completed` sinaliza que você tratou com êxito o evento.</span><span class="sxs-lookup"><span data-stu-id="2262f-114">The call to `event.completed` signals that you have successfully handled the event.</span></span> <span data-ttu-id="2262f-115">Quando uma função é chamada várias vezes, por exemplo, com vários cliques no mesmo comando de suplemento, todos os eventos são enfileirados automaticamente.</span><span class="sxs-lookup"><span data-stu-id="2262f-115">When a function is called multiple times, such as multiple clicks on the same add-in command, all events are automatically queued.</span></span> <span data-ttu-id="2262f-116">O primeiro evento é executado automaticamente, enquanto os outros eventos permanecem na fila.</span><span class="sxs-lookup"><span data-stu-id="2262f-116">The first event runs automatically, while the other events remain on the queue.</span></span> <span data-ttu-id="2262f-117">Quando sua função chama `event.completed`, a próxima chamada em fila para essa função é executada.</span><span class="sxs-lookup"><span data-stu-id="2262f-117">When your function calls `event.completed`, the next queued call to that function runs.</span></span> <span data-ttu-id="2262f-118">Você deve chamar `event.completed`; caso contrário, sua função não será executada.</span><span class="sxs-lookup"><span data-stu-id="2262f-118">You must call `event.completed`; otherwise your function will not run.</span></span>
