---
title: Conceitos fundamentais de programação com a API JavaScript do Word
description: Use as APIs JavaScript do Word para criar suplementos para o Word.
ms.date: 07/28/2020
localization_priority: Priority
ms.openlocfilehash: 1e7a90d4be378ed9b2c1f30ebebd4a0beec45a11
ms.sourcegitcommit: 9609bd5b4982cdaa2ea7637709a78a45835ffb19
ms.translationtype: HT
ms.contentlocale: pt-BR
ms.lasthandoff: 08/28/2020
ms.locfileid: "47293090"
---
# <a name="fundamental-programming-concepts-with-the-word-javascript-api"></a><span data-ttu-id="72bb4-103">Conceitos fundamentais de programação com a API JavaScript do Word</span><span class="sxs-lookup"><span data-stu-id="72bb4-103">Fundamental programming concepts with the Word JavaScript API</span></span>

<span data-ttu-id="72bb4-104">Este artigo descreve conceitos fundamentais para o uso da [API JavaScript do Word](../reference/overview/word-add-ins-reference-overview.md) para criar suplementos para o Word 2016 ou posterior.</span><span class="sxs-lookup"><span data-stu-id="72bb4-104">This article describes concepts that are fundamental to using the [Word JavaScript API](../reference/overview/word-add-ins-reference-overview.md) to build add-ins for Word 2016 or later.</span></span>

## <a name="referencing-officejs"></a><span data-ttu-id="72bb4-105">Referenciando Office.js</span><span class="sxs-lookup"><span data-stu-id="72bb4-105">Referencing Office.js</span></span>

<span data-ttu-id="72bb4-106">Você pode obter referência do Office.js nos seguintes locais:</span><span class="sxs-lookup"><span data-stu-id="72bb4-106">You can reference Office.js from the following locations:</span></span>

- <span data-ttu-id="72bb4-107">`https://appsforoffice.microsoft.com/lib/1/hosted/office.js`: use esse recurso para os suplementos de produção.</span><span class="sxs-lookup"><span data-stu-id="72bb4-107">`https://appsforoffice.microsoft.com/lib/1/hosted/office.js` - use this resource for production add-ins.</span></span>

- <span data-ttu-id="72bb4-108">`https://appsforoffice.microsoft.com/lib/beta/hosted/office.js` - use esse recurso para experimentar recursos de visualização.</span><span class="sxs-lookup"><span data-stu-id="72bb4-108">`https://appsforoffice.microsoft.com/lib/beta/hosted/office.js` - use this resource to try out preview features.</span></span>

## <a name="word-javascript-api-requirement-sets"></a><span data-ttu-id="72bb4-109">Conjuntos de requisitos da API JavaScript do Word</span><span class="sxs-lookup"><span data-stu-id="72bb4-109">Word JavaScript API requirement sets</span></span>

<span data-ttu-id="72bb4-110">Os conjuntos de requisitos são grupos nomeados de membros da API.</span><span class="sxs-lookup"><span data-stu-id="72bb4-110">Requirement sets are named groups of API members.</span></span> <span data-ttu-id="72bb4-111">Os Suplementos do Office usam conjuntos de requisitos especificados no manifesto ou usam uma verificação de tempo de execução para determinar se um aplicativo do Office oferece suporte para as APIs necessárias para um suplemento.</span><span class="sxs-lookup"><span data-stu-id="72bb4-111">Office Add-ins use requirement sets specified in the manifest or use a runtime check to determine whether an Office application supports APIs that an add-in needs.</span></span> <span data-ttu-id="72bb4-112">Para saber mais sobre conjuntos de requisitos da API JavaScript do Word, consulte conjuntos de requisitos da [API JavaScript do Word](../reference/requirement-sets/word-api-requirement-sets.md).</span><span class="sxs-lookup"><span data-stu-id="72bb4-112">For detailed information about Word JavaScript API requirement sets, see [Word JavaScript API requirement sets](../reference/requirement-sets/word-api-requirement-sets.md).</span></span>

## <a name="running-word-add-ins"></a><span data-ttu-id="72bb4-113">Execução de suplementos do Word</span><span class="sxs-lookup"><span data-stu-id="72bb4-113">Running Word add-ins</span></span>

<span data-ttu-id="72bb4-114">Para executar seu suplemento, use um manipulador de eventos `Office.initialize`.</span><span class="sxs-lookup"><span data-stu-id="72bb4-114">To run your add-in, use an `Office.initialize` event handler.</span></span> <span data-ttu-id="72bb4-115">Confira [Entendendo a API](../develop/understanding-the-javascript-api-for-office.md) para saber mais sobre a inicialização do suplemento.</span><span class="sxs-lookup"><span data-stu-id="72bb4-115">For more information about add-in initialization, see [Understanding the API](../develop/understanding-the-javascript-api-for-office.md).</span></span>

<span data-ttu-id="72bb4-116">Os suplementos que visam o Word 2016 ou posterior podem usar as APIs específicas do Word.</span><span class="sxs-lookup"><span data-stu-id="72bb4-116">Add-ins that target Word 2016 or later can use the Word-specific APIs.</span></span> <span data-ttu-id="72bb4-117">Eles passam a lógica de interação do Word como uma função no método `Word.run()`.</span><span class="sxs-lookup"><span data-stu-id="72bb4-117">They pass the Word-interaction logic as a function into the `Word.run()` method.</span></span> <span data-ttu-id="72bb4-118">Confira [Usando o modelo de API específico do aplicativo](../develop/application-specific-api-model.md) para saber mais sobre como interagir com o documento do Word neste modelo de programação.</span><span class="sxs-lookup"><span data-stu-id="72bb4-118">See [Using the application-specific API model](../develop/application-specific-api-model.md) to learn about how to interact with the Word document in this programming model.</span></span>

<span data-ttu-id="72bb4-119">O exemplo a seguir mostra como inicializar e executar um suplemento do Word usando o método `Word.run()`.</span><span class="sxs-lookup"><span data-stu-id="72bb4-119">The following example shows how to initialize and run a Word add-in by using the `Word.run()` method.</span></span>

```js
(function () {
    "use strict";

    // The initialize event handler must be run on each page to initialize Office JS.
    // You can add optional custom initialization code that will run after OfficeJS
    // has initialized.
    Office.initialize = function (reason) {
        // The reason object tells how the add-in was initialized. The values can be:
        // inserted - the add-in was inserted to an open document.
        // documentOpened - the add-in was already inserted in to the document and the document was opened.

        // Checks for the DOM to load using the jQuery ready function.
        $(document).ready(function () {
            // Set your optional initialization code.
            // You can also load saved settings from the Office object.
        });
    };

    // Run a batch operation against the Word JavaScript API object model.
    // Use the context argument to get access to the Word document.
    Word.run(function (context) {

        // Create a proxy object for the document.
        var thisDocument = context.document;
        // ...
    })
})();
```

## <a name="see-also"></a><span data-ttu-id="72bb4-120">Confira também</span><span class="sxs-lookup"><span data-stu-id="72bb4-120">See also</span></span>

- [<span data-ttu-id="72bb4-121">Visão geral da API JavaScript do Word</span><span class="sxs-lookup"><span data-stu-id="72bb4-121">Word JavaScript API overview</span></span>](../reference/overview/word-add-ins-reference-overview.md)
- [<span data-ttu-id="72bb4-122">Criar seu primeiro suplemento do Word</span><span class="sxs-lookup"><span data-stu-id="72bb4-122">Build your first Word add-in</span></span>](../quickstarts/word-quickstart.md)
- [<span data-ttu-id="72bb4-123">Tutorial de suplemento do Word</span><span class="sxs-lookup"><span data-stu-id="72bb4-123">Word add-in tutorial</span></span>](../tutorials/word-tutorial.md)
- [<span data-ttu-id="72bb4-124">Referências da API JavaScript do Word</span><span class="sxs-lookup"><span data-stu-id="72bb4-124">Word JavaScript API reference</span></span>](/javascript/api/word)
