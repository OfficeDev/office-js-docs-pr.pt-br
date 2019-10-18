---
title: Conceitos fundamentais de programação com a API JavaScript do Word
description: Use as APIs JavaScript do Word para criar suplementos para o Word.
ms.date: 07/05/2019
localization_priority: Priority
ms.openlocfilehash: 00a7405d4d89279049d2724dda4fa1384a88dca4
ms.sourcegitcommit: c3673cc693fa7070e1b397922bd735ba3f9342f3
ms.translationtype: HT
ms.contentlocale: pt-BR
ms.lasthandoff: 07/05/2019
ms.locfileid: "35576725"
---
# <a name="fundamental-programming-concepts-with-the-word-javascript-api"></a><span data-ttu-id="5153b-103">Conceitos fundamentais de programação com a API JavaScript do Word</span><span class="sxs-lookup"><span data-stu-id="5153b-103">Fundamental programming concepts with the Excel JavaScript API</span></span>

<span data-ttu-id="5153b-104">Este artigo descreve conceitos fundamentais para o uso da [API JavaScript do Word](../reference/overview/word-add-ins-reference-overview.md) para criar suplementos para o Word 2016 ou posterior.</span><span class="sxs-lookup"><span data-stu-id="5153b-104">This article describes concepts that are fundamental to using the [Word JavaScript API](../reference/overview/word-add-ins-reference-overview.md) to build add-ins for Word 2016 or later.</span></span>

## <a name="referencing-officejs"></a><span data-ttu-id="5153b-105">Referenciando Office.js</span><span class="sxs-lookup"><span data-stu-id="5153b-105">Referencing Office.js</span></span>

<span data-ttu-id="5153b-106">Você pode obter referência do Office.js nos seguintes locais:</span><span class="sxs-lookup"><span data-stu-id="5153b-106">You can reference Office.js from the following locations:</span></span>

- <span data-ttu-id="5153b-107">`https://appsforoffice.microsoft.com/lib/1/hosted/office.js`: use esse recurso para os suplementos de produção.</span><span class="sxs-lookup"><span data-stu-id="5153b-107">`https://appsforoffice.microsoft.com/lib/1/hosted/office.js` - use this resource for production add-ins.</span></span>

- <span data-ttu-id="5153b-108">`https://appsforoffice.microsoft.com/lib/beta/hosted/office.js` - use esse recurso para experimentar recursos de visualização.</span><span class="sxs-lookup"><span data-stu-id="5153b-108">`https://appsforoffice.microsoft.com/lib/beta/hosted/office.js` - use this resource when you're trying out preview features.</span></span>

## <a name="word-javascript-api-requirement-sets"></a><span data-ttu-id="5153b-109">Conjuntos de requisitos da API JavaScript do Word</span><span class="sxs-lookup"><span data-stu-id="5153b-109">Word JavaScript API requirement sets</span></span>

<span data-ttu-id="5153b-110">Os conjuntos de requisitos são grupos nomeados de membros da API.</span><span class="sxs-lookup"><span data-stu-id="5153b-110">Requirement sets are named groups of API members.</span></span> <span data-ttu-id="5153b-111">Os suplementos do Office usam conjuntos de requisitos especificados no manifesto ou usam uma verificação de tempo de execução para determinar se um host do Office oferece suporte para as APIs necessárias para um suplemento.</span><span class="sxs-lookup"><span data-stu-id="5153b-111">Office Add-ins use requirement sets specified in the manifest or use a runtime check to determine whether an Office host supports APIs that an add-in needs.</span></span> <span data-ttu-id="5153b-112">Para saber mais sobre conjuntos de requisitos da API JavaScript do Word, consulte conjuntos de requisitos da [API JavaScript do Word](../reference/requirement-sets/word-api-requirement-sets.md).</span><span class="sxs-lookup"><span data-stu-id="5153b-112">For detailed information about Word JavaScript API requirement sets, see the [Word JavaScript API requirement sets](../reference/requirement-sets/word-api-requirement-sets.md) article.</span></span>

## <a name="running-word-add-ins"></a><span data-ttu-id="5153b-113">Execução de suplementos do Word</span><span class="sxs-lookup"><span data-stu-id="5153b-113">Running Word add-ins</span></span>

<span data-ttu-id="5153b-114">Para executar o suplemento, use um manipulador de eventos **Office.initialize**.</span><span class="sxs-lookup"><span data-stu-id="5153b-114">To run your add-in, use an Office.initialize event handler.</span></span> <span data-ttu-id="5153b-115">Confira [Entendendo a API](/office/dev/add-ins/develop/understanding-the-javascript-api-for-office) para saber mais sobre a inicialização do suplemento.</span><span class="sxs-lookup"><span data-stu-id="5153b-115">For more information about add-in initialization, see [Understanding the API](/office/dev/add-ins/develop/understanding-the-javascript-api-for-office) .</span></span>

<span data-ttu-id="5153b-116">Os suplementos direcionados ao Word 2016 ou posterior são executados passando uma função para o método **Word.run()**.</span><span class="sxs-lookup"><span data-stu-id="5153b-116">Add-ins that target Word 2016 or later execute by passing a function into the **Word.run()** method.</span></span> <span data-ttu-id="5153b-117">A função passada para o método **run** deve ter um argumento de contexto.</span><span class="sxs-lookup"><span data-stu-id="5153b-117">The function passed into the **run** method must have a context argument.</span></span> <span data-ttu-id="5153b-118">Esse [objeto de contexto](/javascript/api/word/word.requestcontext) é diferente do objeto de contexto obtido do objeto do Office, mas ele é usado para interagir com o ambiente de tempo de execução do Word.</span><span class="sxs-lookup"><span data-stu-id="5153b-118">This [context object](/javascript/api/word/word.requestcontext) is different than the context object you get from the Office object, but it is also used to interact with the Word runtime environment.</span></span> <span data-ttu-id="5153b-119">O objeto de contexto fornece acesso ao modelo de objeto da API JavaScript do Word.</span><span class="sxs-lookup"><span data-stu-id="5153b-119">The context object provides access to the Word JavaScript API object model.</span></span> <span data-ttu-id="5153b-120">O exemplo a seguir mostra como iniciar e executar um suplemento do Word usando o método **Word.run()**.</span><span class="sxs-lookup"><span data-stu-id="5153b-120">The following example shows how to initialize and execute a Word add-in by using the **Word.run()** method.</span></span>

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

### <a name="asynchronous-nature-of-word-apis"></a><span data-ttu-id="5153b-121">Natureza assíncrona das APIs do Word</span><span class="sxs-lookup"><span data-stu-id="5153b-121">Asynchronous nature of Excel APIs</span></span>

<span data-ttu-id="5153b-122">A API JavaScript do Word é carregada pelo Office.js.</span><span class="sxs-lookup"><span data-stu-id="5153b-122">The Word JavaScript API is loaded by Office.js.</span></span> <span data-ttu-id="5153b-123">Ela muda a maneira de interagir com objetos, como documentos e parágrafos.</span><span class="sxs-lookup"><span data-stu-id="5153b-123">The Word JavaScript API changes the way that you can interact with objects like documents and paragraphs.</span></span> <span data-ttu-id="5153b-124">Em vez de fornecer APIs assíncronas individuais para recuperar e atualizar cada um desses objetos, a API JavaScript do Word fornece objetos JavaScript “proxy” que correspondem aos objetos reais em execução no Word.</span><span class="sxs-lookup"><span data-stu-id="5153b-124">Rather than providing individual asynchronous APIs for retrieving and updating each of these objects, the Word JavaScript API provides “proxy” JavaScript objects that correspond to the real objects running in Word.</span></span> <span data-ttu-id="5153b-125">Você pode interagir com esses objetos proxy ao ler e gravar, simultaneamente, suas propriedades e chamar, de forma simultânea, métodos para executar operações neles.</span><span class="sxs-lookup"><span data-stu-id="5153b-125">You can interact with these proxy objects by synchronously reading and writing their properties and calling synchronous methods to perform operations on them.</span></span> <span data-ttu-id="5153b-126">Essas interações com objetos proxy não são percebidas imediatamente no script em execução.</span><span class="sxs-lookup"><span data-stu-id="5153b-126">These interactions with proxy objects aren’t immediately realized in the running script.</span></span> <span data-ttu-id="5153b-127">O método **context.sync** sincroniza o estado entre o JavaScript em execução e os objetos reais do Office, executando instruções na fila e recuperando propriedades de objetos carregados do Word para uso no seu script.</span><span class="sxs-lookup"><span data-stu-id="5153b-127">The **context.sync** method synchronizes the state between your running JavaScript and the real objects in Office by executing queued instructions and retrieving properties of loaded Word objects for use in your script.</span></span>

### <a name="synchronizing-word-documents-with-word-javascript-api-proxy-objects"></a><span data-ttu-id="5153b-128">Sincronizar documentos do Word com objetos proxy da API JavaScript do Word</span><span class="sxs-lookup"><span data-stu-id="5153b-128">Synchronizing Word documents with Word JavaScript API proxy objects</span></span>

<span data-ttu-id="5153b-p105">O modelo de objeto da API JavaScript do Word é combinado livremente com os objetos no Word. Os objetos da API JavaScript do Word são proxies de objetos em um documento do Word. As ações executadas em objetos proxy não são percebidas no Word até que o estado do documento seja sincronizado. Por outro lado, o estado do documento do Word não é percebido em objetos proxy, até que o estado do documento seja sincronizado. Para sincronizar o estado do documento, execute o método **context.sync()**. O exemplo a seguir mostra a criação de um objeto proxy do corpo e um comando na fila para carregar a propriedade de texto nesse objeto e usa o método **context.sync()** para sincronizar o corpo do documento do Word com o objeto proxy do corpo.</span><span class="sxs-lookup"><span data-stu-id="5153b-p105">The Word JavaScript API object model is loosely coupled with the objects in Word. Word JavaScript API objects are proxies for objects in a Word document. Actions taken on proxy objects are not realized in Word until the document state has been synchronized. Conversely, the state of the Word document is not realized in the proxy objects until the document state has been synchronized. To synchronize the document state, you run the **context.sync()** method. The following example creates a proxy body object and a queued command to load the text property on the proxy body object, and uses the **context.sync()** method to synchronize the body of the Word document with the body proxy object.</span></span>

```js
// Run a batch operation against the Word object model.
Word.run(function (context) {

    // Create a proxy object for the document body.
    // The body object hasn't been set with any property values.
    var body = context.document.body;

    // Queue a command to load the text property for the proxy document body object.
    body.load("text");

    // Synchronize the document state by executing the queued commands,
    // and return a promise to indicate task completion.
    return context.sync().then(function () {
        console.log("Body contents: " + body.text);
    });
})
```

### <a name="executing-a-batch-of-commands"></a><span data-ttu-id="5153b-135">Execução de um lote de comandos</span><span class="sxs-lookup"><span data-stu-id="5153b-135">Executing a batch of commands</span></span>

<span data-ttu-id="5153b-136">Os objetos proxy do Word têm métodos para acessar e atualizar o modelo de objeto.</span><span class="sxs-lookup"><span data-stu-id="5153b-136">The Word proxy objects have methods for accessing and updating the object model.</span></span> <span data-ttu-id="5153b-137">Esses métodos são executados sequencialmente na ordem em que foram enfileirados no lote.</span><span class="sxs-lookup"><span data-stu-id="5153b-137">These methods are executed sequentially in the order in which they were queued in the batch.</span></span> <span data-ttu-id="5153b-138">Todos os comandos na fila do lote são executados quando o método **context.sync()** é chamado.</span><span class="sxs-lookup"><span data-stu-id="5153b-138">All of the commands that are queued in the batch are executed when context.sync() is called.</span></span>

<span data-ttu-id="5153b-139">O exemplo a seguir mostra como a fila de comandos funciona.</span><span class="sxs-lookup"><span data-stu-id="5153b-139">The following example shows how the command queue works.</span></span> <span data-ttu-id="5153b-140">Quando o método **context.sync()** é chamado, o comando para carregar o corpo de texto é executado no Word.</span><span class="sxs-lookup"><span data-stu-id="5153b-140">When **context.sync()** is called, the command to load the body text is executed in Word.</span></span> <span data-ttu-id="5153b-141">Em seguida, ocorre o comando para inserir o texto no corpo do Word.</span><span class="sxs-lookup"><span data-stu-id="5153b-141">Then, the command to insert text into the body in Word occurs.</span></span> <span data-ttu-id="5153b-142">Na sequência, os resultados são retornados ao objeto proxy do corpo.</span><span class="sxs-lookup"><span data-stu-id="5153b-142">The results are then returned to the body proxy object.</span></span> <span data-ttu-id="5153b-143">O valor da propriedade **body.text**, na API JavaScript do Word, é o valor do corpo do documento do Word, <u>antes</u> da inserção do texto no documento do Word.</span><span class="sxs-lookup"><span data-stu-id="5153b-143">The value of the **body.text** property in the Word JavaScript API is the value of the Word document body <u>before</u> the text was inserted into Word document.</span></span>

```js
// Run a batch operation against the Word JavaScript API.
Word.run(function (context) {

    // Create a proxy object for the document body.
    var body = context.document.body;

    // Queue a command to load the text property of the proxy body object.
    body.load("text");

    // Queue a command to insert text into the end of the Word document body.
    body.insertText('This is text inserted after loading the body.text property',
                    Word.InsertLocation.end);

    // Synchronize the document state by executing the queued commands,
    // and return a promise to indicate task completion.
    return context.sync().then(function () {
        console.log("Body contents: " + body.text);
    });
})
```

## <a name="see-also"></a><span data-ttu-id="5153b-144">Confira também</span><span class="sxs-lookup"><span data-stu-id="5153b-144">See also</span></span>

- [<span data-ttu-id="5153b-145">Visão geral da API JavaScript do Word</span><span class="sxs-lookup"><span data-stu-id="5153b-145">Word JavaScript API overview</span></span>](../reference/overview/word-add-ins-reference-overview.md)
- [<span data-ttu-id="5153b-146">Criar seu primeiro suplemento do Word</span><span class="sxs-lookup"><span data-stu-id="5153b-146">Build your first Word add-in</span></span>](../quickstarts/word-quickstart.md)
- [<span data-ttu-id="5153b-147">Tutorial de suplemento do Word</span><span class="sxs-lookup"><span data-stu-id="5153b-147">Word add-in tutorial</span></span>](../tutorials/word-tutorial.md)
- [<span data-ttu-id="5153b-148">Referências da API JavaScript do Word</span><span class="sxs-lookup"><span data-stu-id="5153b-148">Word JavaScript API reference</span></span>](/javascript/api/word) 


