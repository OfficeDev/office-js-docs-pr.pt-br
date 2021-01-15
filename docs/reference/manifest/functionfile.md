---
title: Elemento FunctionFile no arquivo de manifesto
description: Especifica o arquivo de código-fonte para operações expostas por um suplemento através de comandos de suplemento que executam uma função JavaScript, em vez de exibir a interface do usuário.
ms.date: 11/06/2020
localization_priority: Normal
ms.openlocfilehash: 44bfd514025b8a23f4f6acdf3fec004485ca4c5a
ms.sourcegitcommit: 2f75a37de349251bc0e0fc402c5ae6dc5c3b8b08
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 01/06/2021
ms.locfileid: "49771390"
---
# <a name="functionfile-element"></a><span data-ttu-id="b74b5-103">Elemento FunctionFile</span><span class="sxs-lookup"><span data-stu-id="b74b5-103">FunctionFile element</span></span>

<span data-ttu-id="b74b5-104">Especifica o arquivo de código-fonte para as operações que um suplemento expõe de uma das seguintes maneiras:</span><span class="sxs-lookup"><span data-stu-id="b74b5-104">Specifies the source code file for operations that an add-in exposes in one of the following ways:</span></span>

* <span data-ttu-id="b74b5-105">Comandos de suplemento que executam uma função JavaScript, em vez de exibir a interface do usuário.</span><span class="sxs-lookup"><span data-stu-id="b74b5-105">Add-in commands that execute a JavaScript function instead of displaying UI.</span></span>
* <span data-ttu-id="b74b5-106">Atalhos de teclado que executam uma função JavaScript.</span><span class="sxs-lookup"><span data-stu-id="b74b5-106">Keyboard shortcuts that execute a JavaScript function.</span></span>

<span data-ttu-id="b74b5-107">O `FunctionFile` elemento é um elemento filho de [DesktopFormFactor](desktopformfactor.md) ou [MobileFormFactor](mobileformfactor.md).</span><span class="sxs-lookup"><span data-stu-id="b74b5-107">The `FunctionFile` element is a child element of [DesktopFormFactor](desktopformfactor.md) or [MobileFormFactor](mobileformfactor.md).</span></span> <span data-ttu-id="b74b5-108">O `resid` atributo do `FunctionFile` elemento não pode ter mais de 32 caracteres e é definido como o valor do `id` atributo de um `Url` elemento no `Resources` elemento que contém a URL para um arquivo HTML que contém ou carrega todas as funções JavaScript usadas por botões de comando de suplemento sem interface do usuário, conforme definido pelo [elemento Control](control.md).</span><span class="sxs-lookup"><span data-stu-id="b74b5-108">The `resid` attribute of the `FunctionFile` element can be no more than 32 characters and is set to the value of the `id` attribute of a `Url` element in the `Resources` element that contains the URL to an HTML file that contains or loads all the JavaScript functions used by UI-less add-in command buttons, as defined by the [Control element](control.md).</span></span>

<span data-ttu-id="b74b5-109">Veja a seguir um exemplo do `FunctionFile` elemento.</span><span class="sxs-lookup"><span data-stu-id="b74b5-109">The following is an example of the `FunctionFile` element.</span></span>

```XML
<DesktopFormFactor>
  <FunctionFile resid="residDesktopFuncUrl" />
  <ExtensionPoint xsi:type="PrimaryCommandSurface">
    <!-- information about this extension point -->
  </ExtensionPoint>

  <!-- You can define more than one ExtensionPoint element as needed -->

</DesktopFormFactor>
```

<span data-ttu-id="b74b5-110">O JavaScript no arquivo HTML indicado pelo `FunctionFile` elemento deve chamar `Office.initialize` e definir funções nomeadas que usam um único parâmetro: `event` .</span><span class="sxs-lookup"><span data-stu-id="b74b5-110">The JavaScript in the HTML file indicated by the `FunctionFile` element must call `Office.initialize` and define named functions that take a single parameter: `event`.</span></span> <span data-ttu-id="b74b5-111">As funções devem usar a API `item.notificationMessages` para indicar o progresso, sucesso ou falha ao usuário.</span><span class="sxs-lookup"><span data-stu-id="b74b5-111">The functions should use the `item.notificationMessages` API to indicate progress, success, or failure to the user.</span></span> <span data-ttu-id="b74b5-112">Também deverá chamar `event.completed` quando terminar a execução.</span><span class="sxs-lookup"><span data-stu-id="b74b5-112">It should also call `event.completed` when it has finished execution.</span></span> <span data-ttu-id="b74b5-113">O nome das funções são usados no `FunctionName` elemento para botões sem interface do usuário.</span><span class="sxs-lookup"><span data-stu-id="b74b5-113">The name of the functions are used in the `FunctionName` element for UI-less buttons.</span></span>

<span data-ttu-id="b74b5-114">Veja a seguir um exemplo de um arquivo HTML que define uma `trackMessage` função.</span><span class="sxs-lookup"><span data-stu-id="b74b5-114">The following is an example of an HTML file defining a `trackMessage` function.</span></span>

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

<span data-ttu-id="b74b5-115">O código a seguir mostra como implementar a função usada pelo `FunctionName` .</span><span class="sxs-lookup"><span data-stu-id="b74b5-115">The following code shows how to implement the function used by `FunctionName`.</span></span>

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
> <span data-ttu-id="b74b5-116">A chamada para `event.completed` sinaliza que você tratou com êxito o evento.</span><span class="sxs-lookup"><span data-stu-id="b74b5-116">The call to `event.completed` signals that you have successfully handled the event.</span></span> <span data-ttu-id="b74b5-117">Quando uma função é chamada várias vezes, por exemplo, com vários cliques no mesmo comando de suplemento, todos os eventos são enfileirados automaticamente.</span><span class="sxs-lookup"><span data-stu-id="b74b5-117">When a function is called multiple times, such as multiple clicks on the same add-in command, all events are automatically queued.</span></span> <span data-ttu-id="b74b5-118">O primeiro evento é executado automaticamente, enquanto os outros eventos permanecem na fila.</span><span class="sxs-lookup"><span data-stu-id="b74b5-118">The first event runs automatically, while the other events remain on the queue.</span></span> <span data-ttu-id="b74b5-119">Quando sua função chama `event.completed` , a próxima chamada em fila para essa função é executada.</span><span class="sxs-lookup"><span data-stu-id="b74b5-119">When your function calls `event.completed`, the next queued call to that function runs.</span></span> <span data-ttu-id="b74b5-120">Você deve chamar `event.completed` ; caso contrário, sua função não será executada.</span><span class="sxs-lookup"><span data-stu-id="b74b5-120">You must call `event.completed`; otherwise your function will not run.</span></span>
