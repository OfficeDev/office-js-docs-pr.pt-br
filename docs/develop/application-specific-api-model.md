---
title: Usando o modelo de API específica do aplicativo
description: Saiba mais sobre o modelo de API baseado em promessa para os suplementos do Excel, OneNote e Word.
ms.date: 09/08/2020
localization_priority: Normal
ms.openlocfilehash: fb25201174dcd97b40ccf6be69b238951103db07
ms.sourcegitcommit: c6308cf245ac1bc66a876eaa0a7bb4a2492991ac
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 09/08/2020
ms.locfileid: "47408597"
---
# <a name="using-the-application-specific-api-model"></a><span data-ttu-id="12232-103">Usando o modelo de API específica do aplicativo</span><span class="sxs-lookup"><span data-stu-id="12232-103">Using the application-specific API model</span></span>

<span data-ttu-id="12232-104">Este artigo descreve como usar o modelo de API para criar suplementos no Excel, no Word e no OneNote.</span><span class="sxs-lookup"><span data-stu-id="12232-104">This article describes how to use the API model for building add-ins in Excel, Word, and OneNote.</span></span> <span data-ttu-id="12232-105">Ele apresenta os principais conceitos fundamentais para o uso das APIs baseadas em promessa.</span><span class="sxs-lookup"><span data-stu-id="12232-105">It introduces core concepts that are fundamental to using the promise-based APIs.</span></span>

> [!NOTE]
> <span data-ttu-id="12232-106">Não há suporte para esse modelo nos clientes do Office 2013.</span><span class="sxs-lookup"><span data-stu-id="12232-106">This model is not supported by Office 2013 clients.</span></span> <span data-ttu-id="12232-107">Use o [modelo de API comum](office-javascript-api-object-model.md) para trabalhar com essas versões do Office.</span><span class="sxs-lookup"><span data-stu-id="12232-107">Use the [Common API model](office-javascript-api-object-model.md) to work with those Office versions.</span></span> <span data-ttu-id="12232-108">Para ver as notas de disponibilidade completa da plataforma, confira [disponibilidade de aplicativos e plataformas do cliente Office para suplementos do Office](../overview/office-add-in-availability.md).</span><span class="sxs-lookup"><span data-stu-id="12232-108">For full platform availability notes, see [Office client application and platform availability for Office Add-ins](../overview/office-add-in-availability.md).</span></span>

> [!TIP]
> <span data-ttu-id="12232-109">Os exemplos nesta página usam as APIs JavaScript do Excel, mas os conceitos também se aplicam ao OneNote, Visio e APIs JavaScript do Word.</span><span class="sxs-lookup"><span data-stu-id="12232-109">The examples in this page use the Excel JavaScript APIs, but the concepts also apply to OneNote, Visio, and Word JavaScript APIs.</span></span>

## <a name="asynchronous-nature-of-the-promise-based-apis"></a><span data-ttu-id="12232-110">Natureza assíncrona das APIs baseadas em promessa</span><span class="sxs-lookup"><span data-stu-id="12232-110">Asynchronous nature of the promise-based APIs</span></span>

<span data-ttu-id="12232-111">Os suplementos do Office são sites que aparecem dentro de um contêiner de navegadores em aplicativos do Office, como o Excel.</span><span class="sxs-lookup"><span data-stu-id="12232-111">Office Add-ins are websites which appear inside a browser container within Office applications, such as Excel.</span></span> <span data-ttu-id="12232-112">Esse contêiner é incorporado no aplicativo do Office em plataformas baseadas em área de trabalho, como o Office no Windows, e é executado dentro de um iFrame HTML no Office na Web.</span><span class="sxs-lookup"><span data-stu-id="12232-112">This container is embedded within the Office application on desktop-based platforms, such as Office on Windows, and runs inside an HTML iFrame in Office on the web.</span></span> <span data-ttu-id="12232-113">Devido a considerações de desempenho, as APIs do Office.js não podem interagir de forma síncrona com os aplicativos do Office em todas as plataformas.</span><span class="sxs-lookup"><span data-stu-id="12232-113">Due to performance considerations, the Office.js APIs cannot interact synchronously with the Office applications across all platforms.</span></span> <span data-ttu-id="12232-114">Portanto, a `sync()` chamada de API no Office.js retorna uma [promessa](https://developer.mozilla.org/docs/Web/JavaScript/Reference/Global_Objects/Promise) resolvida quando o aplicativo do Office conclui as ações de leitura ou gravação solicitadas.</span><span class="sxs-lookup"><span data-stu-id="12232-114">Therefore, the `sync()` API call in Office.js returns a [Promise](https://developer.mozilla.org/docs/Web/JavaScript/Reference/Global_Objects/Promise) that is resolved when the Office application completes the requested read or write actions.</span></span> <span data-ttu-id="12232-115">Além disso, você pode enfileirar várias ações, como definir propriedades ou invocar métodos, e executá-las como um lote de comandos com uma única chamada para `sync()` , em vez de enviar uma solicitação separada para cada ação.</span><span class="sxs-lookup"><span data-stu-id="12232-115">Also, you can queue up multiple actions, such as setting properties or invoking methods, and run them as a batch of commands with a single call to `sync()`, rather than sending a separate request for each action.</span></span> <span data-ttu-id="12232-116">As seções a seguir descrevem como fazer isso usando as `run()` `sync()` APIs e.</span><span class="sxs-lookup"><span data-stu-id="12232-116">The following sections describe how to accomplish this using the `run()` and `sync()` APIs.</span></span>

## <a name="run-function"></a><span data-ttu-id="12232-117">função \*. Run</span><span class="sxs-lookup"><span data-stu-id="12232-117">\*.run function</span></span>

<span data-ttu-id="12232-118">`Excel.run`, `Word.run` e `OneNote.run` Execute uma função que especifica as ações a serem executadas em relação ao Excel, Word e OneNote.</span><span class="sxs-lookup"><span data-stu-id="12232-118">`Excel.run`, `Word.run`, and `OneNote.run` execute a function that specifies the actions to perform against Excel, Word, and OneNote.</span></span> <span data-ttu-id="12232-119">`*.run` cria automaticamente um contexto de solicitação que você pode usar para interagir com objetos do Office.</span><span class="sxs-lookup"><span data-stu-id="12232-119">`*.run` automatically creates a request context that you can use to interact with Office objects.</span></span> <span data-ttu-id="12232-120">Quando `*.run` é concluído, uma promessa é resolvida e todos os objetos que foram alocados no tempo de execução são automaticamente liberados.</span><span class="sxs-lookup"><span data-stu-id="12232-120">When `*.run` completes, a promise is resolved, and any objects that were allocated at runtime are automatically released.</span></span>

<span data-ttu-id="12232-121">O exemplo a seguir mostra como usar o `Excel.run` .</span><span class="sxs-lookup"><span data-stu-id="12232-121">The following example shows how to use `Excel.run`.</span></span> <span data-ttu-id="12232-122">O mesmo padrão também é usado com o Word e o OneNote.</span><span class="sxs-lookup"><span data-stu-id="12232-122">The same pattern is also used with Word and OneNote.</span></span>

```js
Excel.run(function (context) {
    // Add your Excel JS API calls here that will be batched and sent to the workbook.
    console.log('Your code goes here.');
}).catch(function (error) {
    // Catch and log any errors that occur within `Excel.run`.
    console.log('error: ' + error);
    if (error instanceof OfficeExtension.Error) {
        console.log('Debug info: ' + JSON.stringify(error.debugInfo));
    }
});
```

## <a name="request-context"></a><span data-ttu-id="12232-123">Contexto de solicitação</span><span class="sxs-lookup"><span data-stu-id="12232-123">Request context</span></span>

<span data-ttu-id="12232-124">O aplicativo do Office e seu suplemento são executados em dois processos diferentes.</span><span class="sxs-lookup"><span data-stu-id="12232-124">The Office application and your add-in run in two different processes.</span></span> <span data-ttu-id="12232-125">Como eles usam diferentes ambientes de tempo de execução, os suplementos exigem um `RequestContext` objeto para conectar seu suplemento a objetos no Office, como planilhas, intervalos, parágrafos e tabelas.</span><span class="sxs-lookup"><span data-stu-id="12232-125">Since they use different runtime environments, add-ins require a `RequestContext` object in order to connect your add-in to objects in Office such as worksheets, ranges, paragraphs, and tables.</span></span> <span data-ttu-id="12232-126">Esse `RequestContext` objeto é fornecido como um argumento ao chamar `*.run` .</span><span class="sxs-lookup"><span data-stu-id="12232-126">This `RequestContext` object is provided as an argument when calling `*.run`.</span></span>

## <a name="proxy-objects"></a><span data-ttu-id="12232-127">Objetos proxy</span><span class="sxs-lookup"><span data-stu-id="12232-127">Proxy objects</span></span>

<span data-ttu-id="12232-128">Os objetos JavaScript do Office que você declara e usa com as APIs baseadas em promessa são objetos de proxy.</span><span class="sxs-lookup"><span data-stu-id="12232-128">The Office JavaScript objects that you declare and use with the promise-based APIs are proxy objects.</span></span> <span data-ttu-id="12232-129">Todos os métodos invocados, ou as propriedades definidas ou carregadas em objetos proxy são simplesmente adicionados a uma fila de comandos pendentes.</span><span class="sxs-lookup"><span data-stu-id="12232-129">Any methods that you invoke or properties that you set or load on proxy objects are simply added to a queue of pending commands.</span></span> <span data-ttu-id="12232-130">Quando você chama o `sync()` método no contexto de solicitação (por exemplo, `context.sync()` ), os comandos enfileirados são expedidos para o aplicativo do Office e executados.</span><span class="sxs-lookup"><span data-stu-id="12232-130">When you call the `sync()` method on the request context (for example, `context.sync()`), the queued commands are dispatched to the Office application and run.</span></span> <span data-ttu-id="12232-131">Essas APIs são essencialmente centradas em lote.</span><span class="sxs-lookup"><span data-stu-id="12232-131">These APIs are fundamentally batch-centric.</span></span> <span data-ttu-id="12232-132">Você pode enfileirar quantas alterações desejar no contexto da solicitação e, em seguida, chamar o `sync()` método para executar o lote de comandos enfileirados.</span><span class="sxs-lookup"><span data-stu-id="12232-132">You can queue up as many changes as you wish on the request context, and then call the `sync()` method to run the batch of queued commands.</span></span>

<span data-ttu-id="12232-133">Por exemplo, o trecho de código a seguir declara o objeto JavaScript [Excel. Range](/javascript/api/excel/excel.range) local, `selectedRange` para fazer referência a um intervalo selecionado na pasta de trabalho do Excel e, em seguida, define algumas propriedades nesse objeto.</span><span class="sxs-lookup"><span data-stu-id="12232-133">For example, the following code snippet declares the local JavaScript [Excel.Range](/javascript/api/excel/excel.range) object, `selectedRange`, to reference a selected range in the Excel workbook, and then sets some properties on that object.</span></span> <span data-ttu-id="12232-134">O `selectedRange` objeto é um objeto proxy, portanto, as propriedades que são definidas e o método invocado nesse objeto não serão refletidas no documento do Excel até que seu suplemento chame `context.sync()` .</span><span class="sxs-lookup"><span data-stu-id="12232-134">The `selectedRange` object is a proxy object, so the properties that are set and the method that is invoked on that object will not be reflected in the Excel document until your add-in calls `context.sync()`.</span></span>

```js
var selectedRange = context.workbook.getSelectedRange();
selectedRange.format.fill.color = "#4472C4";
selectedRange.format.font.color = "white";
selectedRange.format.autofitColumns();
```

### <a name="performance-tip-minimize-the-number-of-proxy-objects-created"></a><span data-ttu-id="12232-135">Dica de desempenho: minimizar o número de objetos de proxy criados</span><span class="sxs-lookup"><span data-stu-id="12232-135">Performance tip: Minimize the number of proxy objects created</span></span>

<span data-ttu-id="12232-136">Evite criar repetidamente o mesmo objeto proxy.</span><span class="sxs-lookup"><span data-stu-id="12232-136">Avoid repeatedly creating the same proxy object.</span></span> <span data-ttu-id="12232-137">Em vez disso, se você precisar do mesmo objeto proxy para mais de uma operação, crie-o uma vez e o atribua a uma variável, em seguida, use essa variável no seu código.</span><span class="sxs-lookup"><span data-stu-id="12232-137">Instead, if you need the same proxy object for more than one operation, create it once and assign it to a variable, then use that variable in your code.</span></span>

```js
// BAD: Repeated calls to .getRange() to create the same proxy object.
worksheet.getRange("A1").format.fill.color = "red";
worksheet.getRange("A1").numberFormat = "0.00%";
worksheet.getRange("A1").values = [[1]];

// GOOD: Create the range proxy object once and assign to a variable.
var range = worksheet.getRange("A1")
range.format.fill.color = "red";
range.numberFormat = "0.00%";
range.values = [[1]];

// ALSO GOOD: Use a "set" method to immediately set all the properties without even needing to create a variable!
worksheet.getRange("A1").set({
    numberFormat: [["0.00%"]],
    values: [[1]],
    format: {
        fill: {
            color: "red"
        }
    }
});
```

### <a name="sync"></a><span data-ttu-id="12232-138">sync()</span><span class="sxs-lookup"><span data-stu-id="12232-138">sync()</span></span>

<span data-ttu-id="12232-139">Chamar o `sync()` método no contexto de solicitação sincroniza o estado entre objetos de proxy e objetos no documento do Office.</span><span class="sxs-lookup"><span data-stu-id="12232-139">Calling the `sync()` method on the request context synchronizes the state between proxy objects and objects in the Office document.</span></span> <span data-ttu-id="12232-140">O `sync()` método executa todos os comandos que estão na fila no contexto de solicitação e recupera valores para todas as propriedades que devem ser carregadas nos objetos de proxy.</span><span class="sxs-lookup"><span data-stu-id="12232-140">The `sync()` method runs any commands that are queued on the request context and retrieves values for any properties that should be loaded on the proxy objects.</span></span> <span data-ttu-id="12232-141">O `sync()` método é executado de forma assíncrona e retorna uma [promessa](https://developer.mozilla.org/docs/Web/JavaScript/Reference/Global_Objects/Promise), que é resolvida quando o `sync()` método é concluído.</span><span class="sxs-lookup"><span data-stu-id="12232-141">The `sync()` method executes asynchronously and returns a [Promise](https://developer.mozilla.org/docs/Web/JavaScript/Reference/Global_Objects/Promise), which is resolved when the `sync()` method completes.</span></span>

<span data-ttu-id="12232-142">O exemplo a seguir mostra uma função em lotes que define um objeto de proxy JavaScript local ( `selectedRange` ), carrega uma propriedade desse objeto e, em seguida, usa o padrão de promessas do JavaScript a ser chamado `context.sync()` para sincronizar o estado entre objetos proxy e objetos no documento do Excel.</span><span class="sxs-lookup"><span data-stu-id="12232-142">The following example shows a batch function that defines a local JavaScript proxy object (`selectedRange`), loads a property of that object, and then uses the JavaScript promises pattern to call `context.sync()` to synchronize the state between proxy objects and objects in the Excel document.</span></span>

```js
Excel.run(function (context) {
    var selectedRange = context.workbook.getSelectedRange();
    selectedRange.load('address');
    return context.sync()
      .then(function () {
        console.log('The selected range is: ' + selectedRange.address);
    });
}).catch(function (error) {
    console.log('error: ' + error);
    if (error instanceof OfficeExtension.Error) {
        console.log('Debug info: ' + JSON.stringify(error.debugInfo));
    }
});
```

<span data-ttu-id="12232-143">No exemplo anterior, `selectedRange` é definido e sua propriedade `address` é carregada quando `context.sync()` é chamado.</span><span class="sxs-lookup"><span data-stu-id="12232-143">In the previous example, `selectedRange` is set and its `address` property is loaded when `context.sync()` is called.</span></span>

<span data-ttu-id="12232-144">Como `sync()` é uma operação assíncrona, você sempre deve retornar o `Promise` objeto para garantir que a `sync()` operação seja concluída antes de o script continuar a ser executado.</span><span class="sxs-lookup"><span data-stu-id="12232-144">Since `sync()` is an asynchronous operation, you should always return the `Promise` object to ensure the `sync()` operation completes before the script continues to run.</span></span> <span data-ttu-id="12232-145">Se você estiver usando o TypeScript ou ES6 + JavaScript, você `await` poderá `context.sync()` chamar em vez de retornar a promessa.</span><span class="sxs-lookup"><span data-stu-id="12232-145">If you're using TypeScript or ES6+ JavaScript, you can `await` the `context.sync()` call instead of returning the promise.</span></span>

#### <a name="performance-tip-minimize-the-number-of-sync-calls"></a><span data-ttu-id="12232-146">Dica de desempenho: minimizar o número de chamadas de sincronização</span><span class="sxs-lookup"><span data-stu-id="12232-146">Performance tip: Minimize the number of sync calls</span></span>

<span data-ttu-id="12232-147">Na API do JavaScript do Excel, `sync()` é a única operação assíncrona e pode ser lenta em algumas circunstâncias, especialmente no Excel Online na Web.</span><span class="sxs-lookup"><span data-stu-id="12232-147">In the Excel JavaScript API, `sync()` is the only asynchronous operation, and it can be slow under some circumstances, especially for Excel on the web.</span></span> <span data-ttu-id="12232-148">Para otimizar o desempenho, minimize o número de chamadas para `sync()`, enfileirando o maior número possível de alterações antes de chamá-lo.</span><span class="sxs-lookup"><span data-stu-id="12232-148">To optimize performance, minimize the number of calls to `sync()` by queueing up as many changes as possible before calling it.</span></span> <span data-ttu-id="12232-149">Para obter mais informações sobre como otimizar `sync()` o desempenho do, consulte [Evite usar o método Context. Sync em loops](../concepts/correlated-objects-pattern.md).</span><span class="sxs-lookup"><span data-stu-id="12232-149">For more information about optimizing performance with `sync()`, see [Avoid using the context.sync method in loops](../concepts/correlated-objects-pattern.md).</span></span>

### <a name="load"></a><span data-ttu-id="12232-150">load()</span><span class="sxs-lookup"><span data-stu-id="12232-150">load()</span></span>

<span data-ttu-id="12232-151">Antes de poder ler as propriedades de um objeto proxy, você deve carregar explicitamente as propriedades para preencher o objeto proxy com dados do documento do Office e, em seguida, chamar `context.sync()` .</span><span class="sxs-lookup"><span data-stu-id="12232-151">Before you can read the properties of a proxy object, you must explicitly load the properties to populate the proxy object with data from the Office document, and then call `context.sync()`.</span></span> <span data-ttu-id="12232-152">Por exemplo, se você criar um objeto proxy para fazer referência a um intervalo selecionado e, em seguida, quiser ler a propriedade do intervalo selecionado `address` , você precisará carregar a `address` propriedade antes de poder lê-la.</span><span class="sxs-lookup"><span data-stu-id="12232-152">For example, if you create a proxy object to reference a selected range, and then want to read the selected range's `address` property, you need to load the `address` property before you can read it.</span></span> <span data-ttu-id="12232-153">Para solicitar que as propriedades de um objeto proxy sejam carregadas, chame o `load()` método no objeto e especifique as propriedades a serem carregadas.</span><span class="sxs-lookup"><span data-stu-id="12232-153">To request properties of a proxy object be loaded, call the `load()` method on the object and specify the properties to load.</span></span> <span data-ttu-id="12232-154">O exemplo a seguir mostra a `Range.address` propriedade que está sendo carregada `myRange` .</span><span class="sxs-lookup"><span data-stu-id="12232-154">The following example shows the `Range.address` property being loaded for `myRange`.</span></span>

```js
Excel.run(function (context) {
    var sheetName = 'Sheet1';
    var rangeAddress = 'A1:B2';
    var myRange = context.workbook.worksheets.getItem(sheetName).getRange(rangeAddress);

    myRange.load('address');

    return context.sync()
      .then(function () {
        console.log (myRange.address);   // ok
        //console.log (myRange.values);  // not ok as it was not loaded
        });
    }).then(function () {
        console.log('done');
}).catch(function (error) {
    console.log('Error: ' + error);
    if (error instanceof OfficeExtension.Error) {
        console.log('Debug info: ' + JSON.stringify(error.debugInfo));
    }
});
```

> [!NOTE]
> <span data-ttu-id="12232-155">Se você estiver apenas chamando métodos ou definindo propriedades em um objeto proxy, você não precisa chamar o `load()` método.</span><span class="sxs-lookup"><span data-stu-id="12232-155">If you are only calling methods or setting properties on a proxy object, you don't need to call the `load()` method.</span></span> <span data-ttu-id="12232-156">O `load()` método só é necessário quando você deseja ler propriedades em um objeto proxy.</span><span class="sxs-lookup"><span data-stu-id="12232-156">The `load()` method is only required when you want to read properties on a proxy object.</span></span>

<span data-ttu-id="12232-p115">Assim como as solicitações para definir propriedades ou invocar métodos em objetos proxy, as solicitações para carregar propriedades em objetos proxy são adicionadas à fila de comandos pendentes no contexto de solicitação, sendo executadas na próxima vez que você chamar o método `sync()`. É possível enfileirar quantas chamadas de `load()` forem necessárias no contexto de solicitação.</span><span class="sxs-lookup"><span data-stu-id="12232-p115">Just like requests to set properties or invoke methods on proxy objects, requests to load properties on proxy objects get added to the queue of pending commands on the request context, which will run the next time you call the `sync()` method. You can queue up as many `load()` calls on the request context as necessary.</span></span>

#### <a name="scalar-and-navigation-properties"></a><span data-ttu-id="12232-159">Propriedades escalares e de navegação</span><span class="sxs-lookup"><span data-stu-id="12232-159">Scalar and navigation properties</span></span>

<span data-ttu-id="12232-160">Há duas categorias de propriedades: **escalar** e de **navegação**.</span><span class="sxs-lookup"><span data-stu-id="12232-160">There are two categories of properties: **scalar** and **navigational**.</span></span> <span data-ttu-id="12232-161">As propriedades escalares são tipos atribuíveis, como cadeias de caracteres, inteiros e estruturas JSON.</span><span class="sxs-lookup"><span data-stu-id="12232-161">Scalar properties are assignable types such as strings, integers, and JSON structs.</span></span> <span data-ttu-id="12232-162">As propriedades de navegação são objetos somente leitura e coleções de objetos que têm seus campos atribuídos, em vez de atribuir diretamente a propriedade.</span><span class="sxs-lookup"><span data-stu-id="12232-162">Navigation properties are read-only objects and collections of objects that have their fields assigned, instead of directly assigning the property.</span></span> <span data-ttu-id="12232-163">Por exemplo, `name` e `position` os membros do objeto [Excel. Worksheet](/javascript/api/excel/excel.worksheet) são propriedades escalares, enquanto `protection` e `tables` são propriedades de navegação.</span><span class="sxs-lookup"><span data-stu-id="12232-163">For example, `name` and `position` members on the [Excel.Worksheet](/javascript/api/excel/excel.worksheet) object are scalar properties, whereas `protection` and `tables` are navigation properties.</span></span>

<span data-ttu-id="12232-164">O suplemento pode usar propriedades de navegação como um caminho para carregar Propriedades escalares específicas.</span><span class="sxs-lookup"><span data-stu-id="12232-164">Your add-in can use navigational properties as a path to load specific scalar properties.</span></span> <span data-ttu-id="12232-165">O código a seguir enfileira um `load` comando para o nome da fonte usada por um `Excel.Range` objeto, sem carregar nenhuma outra informação.</span><span class="sxs-lookup"><span data-stu-id="12232-165">The following code queues up a `load` command for the name of the font used by an `Excel.Range` object, without loading any other information.</span></span>

```js
someRange.load("format/font/name")
```

<span data-ttu-id="12232-166">Você também pode definir as propriedades escalares de uma propriedade de navegação atravessando o caminho.</span><span class="sxs-lookup"><span data-stu-id="12232-166">You can also set the scalar properties of a navigation property by traversing the path.</span></span> <span data-ttu-id="12232-167">Por exemplo, você pode definir o tamanho da fonte de um `Excel.Range` usando `someRange.format.font.size = 10;` .</span><span class="sxs-lookup"><span data-stu-id="12232-167">For example, you could set the font size for an `Excel.Range` by using `someRange.format.font.size = 10;`.</span></span> <span data-ttu-id="12232-168">Você não precisa carregar a propriedade antes de defini-la.</span><span class="sxs-lookup"><span data-stu-id="12232-168">You don't need to load the property before you set it.</span></span>

<span data-ttu-id="12232-169">Observe que algumas das propriedades em um objeto podem ter o mesmo nome de outro objeto.</span><span class="sxs-lookup"><span data-stu-id="12232-169">Please be aware that some of the properties under an object may have the same name as another object.</span></span> <span data-ttu-id="12232-170">Por exemplo, `format` é uma propriedade sob o `Excel.Range` objeto, mas `format` também é um objeto.</span><span class="sxs-lookup"><span data-stu-id="12232-170">For example, `format` is a property under the `Excel.Range` object, but `format` itself is an object as well.</span></span> <span data-ttu-id="12232-171">Portanto, se você fizer uma chamada como `range.load("format")` , isso equivale a `range.format.load()` (uma instrução vazia indesejável `load()` ).</span><span class="sxs-lookup"><span data-stu-id="12232-171">So, if you make a call such as `range.load("format")`, this is equivalent to `range.format.load()` (an undesirable empty `load()` statement).</span></span> <span data-ttu-id="12232-172">Para evitar isso, o código só deve carregar os "nós folha" em uma árvore de objetos.</span><span class="sxs-lookup"><span data-stu-id="12232-172">To avoid this, your code should only load the "leaf nodes" in an object tree.</span></span>

#### <a name="calling-load-without-parameters-not-recommended"></a><span data-ttu-id="12232-173">Chamar `load` sem parâmetros (não recomendado)</span><span class="sxs-lookup"><span data-stu-id="12232-173">Calling `load` without parameters (not recommended)</span></span>

<span data-ttu-id="12232-174">Se você chamar o `load()` método em um objeto (ou coleção) sem especificar nenhum parâmetro, todas as propriedades escalares do objeto ou dos objetos da coleção serão carregadas.</span><span class="sxs-lookup"><span data-stu-id="12232-174">If you call the `load()` method on an object (or collection) without specifying any parameters, all scalar properties of the object or the collection's objects will be loaded.</span></span> <span data-ttu-id="12232-175">O carregamento de dados desnecessários tornará o suplemento lento.</span><span class="sxs-lookup"><span data-stu-id="12232-175">Loading unneeded data will slow down your add-in.</span></span> <span data-ttu-id="12232-176">Você sempre deve especificar explicitamente as propriedades a serem carregadas.</span><span class="sxs-lookup"><span data-stu-id="12232-176">You should always explicitly specify which properties to load.</span></span>

> [!IMPORTANT]
> <span data-ttu-id="12232-177">A quantidade de dados retornados por uma declaração `load` sem parâmetros pode exceder os limites de tamanho do serviço.</span><span class="sxs-lookup"><span data-stu-id="12232-177">The amount of data returned by a parameter-less `load` statement can exceed the size limits of the service.</span></span> <span data-ttu-id="12232-178">Para reduzir os riscos a suplementos mais antigos, algumas propriedades não são retornadas por `load` sem a solicitação explícita.</span><span class="sxs-lookup"><span data-stu-id="12232-178">To reduce the risks to older add-ins, some properties are not returned by `load` without explicitly requesting them.</span></span> <span data-ttu-id="12232-179">As seguintes propriedades são excluídas dessas operações de carregamento:</span><span class="sxs-lookup"><span data-stu-id="12232-179">The following properties are excluded from such load operations:</span></span>
>
> * `Excel.Range.numberFormatCategories`

### <a name="clientresult"></a><span data-ttu-id="12232-180">ClientResult</span><span class="sxs-lookup"><span data-stu-id="12232-180">ClientResult</span></span>

<span data-ttu-id="12232-181">Os métodos nas APIs baseadas em promessa que retornam tipos primitivos têm um padrão semelhante ao `load` / `sync` paradigma.</span><span class="sxs-lookup"><span data-stu-id="12232-181">Methods in the promise-based APIs that return primitive types have a similar pattern to the `load`/`sync` paradigm.</span></span> <span data-ttu-id="12232-182">Por exemplo, `Excel.TableCollection.getCount` obtém o número de tabelas da coleção.</span><span class="sxs-lookup"><span data-stu-id="12232-182">As an example, `Excel.TableCollection.getCount` gets the number of tables in the collection.</span></span> <span data-ttu-id="12232-183">`getCount` Retorna um `ClientResult<number>` , significando que a `value` propriedade no retornado [`ClientResult`](/javascript/api/office/officeextension.clientresult) é um número.</span><span class="sxs-lookup"><span data-stu-id="12232-183">`getCount` returns a `ClientResult<number>`, meaning the `value` property in the returned [`ClientResult`](/javascript/api/office/officeextension.clientresult) is a number.</span></span> <span data-ttu-id="12232-184">Seu script não pode acessar esse valor até que `context.sync()` seja chamado.</span><span class="sxs-lookup"><span data-stu-id="12232-184">Your script can't access that value until `context.sync()` is called.</span></span>

<span data-ttu-id="12232-185">O código a seguir obtém o número total de tabelas em uma pasta de trabalho do Excel e registra esse número no console.</span><span class="sxs-lookup"><span data-stu-id="12232-185">The following code gets the total number of tables in an Excel workbook and logs that number to the console.</span></span>

```js
var tableCount = context.workbook.tables.getCount();

// This sync call implicitly loads tableCount.value.
// Any other ClientResult values are loaded too.
return context.sync()
    .then(function () {
        // Trying to log the value before calling sync would throw an error.
        console.log (tableCount.value);
    });
```

### <a name="set"></a><span data-ttu-id="12232-186">set()</span><span class="sxs-lookup"><span data-stu-id="12232-186">set()</span></span>

<span data-ttu-id="12232-187">A definição de propriedades em um objeto com propriedades de navegação aninhadas pode ser uma tarefa complicada.</span><span class="sxs-lookup"><span data-stu-id="12232-187">Setting properties on an object with nested navigation properties can be cumbersome.</span></span> <span data-ttu-id="12232-188">Como alternativa à definição de propriedades individuais usando caminhos de navegação, conforme descrito acima, você pode usar o `object.set()` método que está disponível em objetos nas APIs JavaScript baseadas em promessa.</span><span class="sxs-lookup"><span data-stu-id="12232-188">As an alternative to setting individual properties using navigation paths as described above, you can use the `object.set()` method that is available on objects in the promise-based JavaScript APIs.</span></span> <span data-ttu-id="12232-189">Com esse método, é possível definir várias propriedades de um objeto de uma vez passando outro objeto do mesmo tipo Office.js ou um objeto JavaScript com propriedades que são estruturadas, como as propriedades do objeto no qual o método é chamado.</span><span class="sxs-lookup"><span data-stu-id="12232-189">With this method, you can set multiple properties of an object at once by passing either another object of the same Office.js type or a JavaScript object with properties that are structured like the properties of the object on which the method is called.</span></span>

<span data-ttu-id="12232-p124">O exemplo de código a seguir define várias propriedades do formato de um intervalo chamando o método `set()` e passando um objeto JavaScript com nomes e tipos de propriedade que espelham a estrutura das propriedades no objeto `Range`. Este exemplo supõe que há dados no intervalo **B2:E2**.</span><span class="sxs-lookup"><span data-stu-id="12232-p124">The following code sample sets several format properties of a range by calling the `set()` method and passing in a JavaScript object with property names and types that mirror the structure of properties in the `Range` object. This example assumes that there is data in range **B2:E2**.</span></span>

```js
Excel.run(function (ctx) {
    var sheet = ctx.workbook.worksheets.getItem("Sample");
    var range = sheet.getRange("B2:E2");
    range.set({
        format: {
            fill: {
                color: '#4472C4'
            },
            font: {
                name: 'Verdana',
                color: 'white'
            }
        }
    });
    range.format.autofitColumns();

    return ctx.sync();
}).catch(function(error) {
    console.log("Error: " + error);
    if (error instanceof OfficeExtension.Error) {
        console.log("Debug info: " + JSON.stringify(error.debugInfo));
    }
});
```

### <a name="some-properties-cannot-be-set-directly"></a><span data-ttu-id="12232-192">Algumas propriedades não podem ser definidas diretamente</span><span class="sxs-lookup"><span data-stu-id="12232-192">Some properties cannot be set directly</span></span>

<span data-ttu-id="12232-193">Algumas propriedades não podem ser definidas, apesar de serem graváveis.</span><span class="sxs-lookup"><span data-stu-id="12232-193">Some properties cannot be set, despite being writable.</span></span> <span data-ttu-id="12232-194">Essas propriedades fazem parte de uma propriedade pai que deve ser definida como um único objeto.</span><span class="sxs-lookup"><span data-stu-id="12232-194">These properties are part of a parent property that must be set as a single object.</span></span> <span data-ttu-id="12232-195">Isso ocorre porque a propriedade Parent depende das subpropriedades que têm relações lógicas específicas.</span><span class="sxs-lookup"><span data-stu-id="12232-195">This is because that parent property relies on the subproperties having specific, logical relationships.</span></span> <span data-ttu-id="12232-196">Essas propriedades pai devem ser definidas usando a notação literal de objeto para definir o objeto inteiro, em vez de definir as subpropriedades individuais desse objeto.</span><span class="sxs-lookup"><span data-stu-id="12232-196">These parent properties must be set using object literal notation to set the entire object, instead of setting that object's individual subproperties.</span></span> <span data-ttu-id="12232-197">Um exemplo disso é encontrado no [PageLayout](/javascript/api/excel/excel.pagelayout).</span><span class="sxs-lookup"><span data-stu-id="12232-197">One example of this is found in [PageLayout](/javascript/api/excel/excel.pagelayout).</span></span> <span data-ttu-id="12232-198">A `zoom` propriedade deve ser definida com um único objeto [PageLayoutZoomOptions](/javascript/api/excel/excel.pagelayoutzoomoptions) , conforme mostrado aqui:</span><span class="sxs-lookup"><span data-stu-id="12232-198">The `zoom` property must be set with a single [PageLayoutZoomOptions](/javascript/api/excel/excel.pagelayoutzoomoptions) object, as shown here:</span></span>

```js
// PageLayout.zoom.scale must be set by assigning PageLayout.zoom to a PageLayoutZoomOptions object.
sheet.pageLayout.zoom = { scale: 200 };
```

<span data-ttu-id="12232-199">No exemplo anterior, você ***não*** poderá atribuir `zoom` um valor diretamente: `sheet.pageLayout.zoom.scale = 200;` .</span><span class="sxs-lookup"><span data-stu-id="12232-199">In the previous example, you would ***not*** be able to directly assign `zoom` a value: `sheet.pageLayout.zoom.scale = 200;`.</span></span> <span data-ttu-id="12232-200">Essa instrução gera um erro porque `zoom` não está carregada.</span><span class="sxs-lookup"><span data-stu-id="12232-200">That statement throws an error because `zoom` is not loaded.</span></span> <span data-ttu-id="12232-201">Mesmo que `zoom` fosse carregado, o conjunto de escala não terá efeito.</span><span class="sxs-lookup"><span data-stu-id="12232-201">Even if `zoom` were to be loaded, the set of scale will not take effect.</span></span> <span data-ttu-id="12232-202">Todas as operações de contexto acontecem em `zoom` , atualizando o objeto de proxy no suplemento e substituindo os valores definidos localmente.</span><span class="sxs-lookup"><span data-stu-id="12232-202">All context operations happen on `zoom`, refreshing the proxy object in the add-in and overwriting locally set values.</span></span>

<span data-ttu-id="12232-203">Esse comportamento difere das [Propriedades de navegação](application-specific-api-model.md#scalar-and-navigation-properties) , como [Range. Format](/javascript/api/excel/excel.range#format).</span><span class="sxs-lookup"><span data-stu-id="12232-203">This behavior differs from [navigational properties](application-specific-api-model.md#scalar-and-navigation-properties) like [Range.format](/javascript/api/excel/excel.range#format).</span></span> <span data-ttu-id="12232-204">As propriedades de `format` podem ser definidas usando a navegação de objeto, conforme mostrado aqui:</span><span class="sxs-lookup"><span data-stu-id="12232-204">Properties of `format` can be set using object navigation, as shown here:</span></span>

```js
// This will set the font size on the range during the next `content.sync()`.
range.format.font.size = 10;
```

<span data-ttu-id="12232-205">Você pode identificar uma propriedade que não pode ter suas subpropriedades definidas diretamente verificando seu modificador somente leitura.</span><span class="sxs-lookup"><span data-stu-id="12232-205">You can identify a property that cannot have its subproperties directly set by checking its read-only modifier.</span></span> <span data-ttu-id="12232-206">Todas as propriedades somente leitura podem ter suas subpropriedades não somente leitura definidas diretamente.</span><span class="sxs-lookup"><span data-stu-id="12232-206">All read-only properties can have their non-read-only subproperties directly set.</span></span> <span data-ttu-id="12232-207">Propriedades graváveis como `PageLayout.zoom` devem ser definidas com um objeto nesse nível.</span><span class="sxs-lookup"><span data-stu-id="12232-207">Writeable properties like `PageLayout.zoom` must be set with an object at that level.</span></span> <span data-ttu-id="12232-208">Em Resumo:</span><span class="sxs-lookup"><span data-stu-id="12232-208">In summary:</span></span>

- <span data-ttu-id="12232-209">Propriedade somente leitura: as subpropriedades podem ser definidas por meio de navegação.</span><span class="sxs-lookup"><span data-stu-id="12232-209">Read-only property: Subproperties can be set through navigation.</span></span>
- <span data-ttu-id="12232-210">Propriedade writable: as subpropriedades não podem ser definidas por meio de navegação (devem ser definidas como parte da atribuição de objeto pai inicial).</span><span class="sxs-lookup"><span data-stu-id="12232-210">Writable property: Subproperties cannot be set through navigation (must be set as part of the initial parent object assignment).</span></span>



## <a name="ornullobject-methods-and-properties"></a><span data-ttu-id="12232-211">Métodos e propriedades do &#42;OrNullObject</span><span class="sxs-lookup"><span data-stu-id="12232-211">&#42;OrNullObject methods and properties</span></span>

<span data-ttu-id="12232-212">Alguns métodos e propriedades de assessor geram uma exceção quando o objeto desejado não existe.</span><span class="sxs-lookup"><span data-stu-id="12232-212">Some accessor methods and properties throw an exception when the desired object doesn't exist.</span></span> <span data-ttu-id="12232-213">Por exemplo, se você tentar obter uma planilha do Excel especificando um nome de planilha que não esteja na pasta de trabalho, o `getItem()` método gera uma `ItemNotFound` exceção.</span><span class="sxs-lookup"><span data-stu-id="12232-213">For example, if you attempt to get an Excel worksheet by specifying a worksheet name that isn't in the workbook, the `getItem()` method throws an `ItemNotFound` exception.</span></span> <span data-ttu-id="12232-214">As bibliotecas específicas do aplicativo fornecem uma maneira de seu código testar a existência de entidades de documento sem exigir código de tratamento de exceção.</span><span class="sxs-lookup"><span data-stu-id="12232-214">The application-specific libraries provide a way for your code to test for the existence of document entities without requiring exception handling code.</span></span> <span data-ttu-id="12232-215">Isso é feito usando as `*OrNullObject` variações de métodos e propriedades.</span><span class="sxs-lookup"><span data-stu-id="12232-215">This is accomplished by using the `*OrNullObject` variations of methods and properties.</span></span> <span data-ttu-id="12232-216">Essas variações retornam um objeto cuja `isNullObject` propriedade é definida como `true` , se o item especificado não existir, em vez de gerar uma exceção.</span><span class="sxs-lookup"><span data-stu-id="12232-216">These variations return an object whose `isNullObject` property is set to `true`, if the specified item doesn't exist, rather than throwing an exception.</span></span>

<span data-ttu-id="12232-217">Por exemplo, você pode chamar o `getItemOrNullObject()` método em uma coleção como **planilhas** para recuperar um item da coleção.</span><span class="sxs-lookup"><span data-stu-id="12232-217">For example, you can call the `getItemOrNullObject()` method on a collection such as **Worksheets** to retrieve an item from the collection.</span></span> <span data-ttu-id="12232-218">O `getItemOrNullObject()` método retorna o item especificado se ele existir; caso contrário, retorna um objeto cuja `isNullObject` propriedade está definida como `true` .</span><span class="sxs-lookup"><span data-stu-id="12232-218">The `getItemOrNullObject()` method returns the specified item if it exists; otherwise, it returns an object whose `isNullObject` property is set to `true`.</span></span> <span data-ttu-id="12232-219">Seu código pode então avaliar essa propriedade para determinar se o objeto existe.</span><span class="sxs-lookup"><span data-stu-id="12232-219">Your code can then evaluate this property to determine whether the object exists.</span></span>

> [!NOTE]
> <span data-ttu-id="12232-220">As `*OrNullObject` variações nunca retornam o valor de JavaScript `null` .</span><span class="sxs-lookup"><span data-stu-id="12232-220">The `*OrNullObject` variations do not ever return the JavaScript value `null`.</span></span> <span data-ttu-id="12232-221">Eles retornam objetos de proxy do Office comuns.</span><span class="sxs-lookup"><span data-stu-id="12232-221">They return ordinary Office proxy objects.</span></span> <span data-ttu-id="12232-222">Se a entidade que o objeto representa não existir, a `isNullObject` Propriedade do objeto será definida como `true` .</span><span class="sxs-lookup"><span data-stu-id="12232-222">If the the entity that the object represents does not exist then the `isNullObject` property of the object is set to `true`.</span></span> <span data-ttu-id="12232-223">Não teste o objeto retornado para nulidade ou falsity.</span><span class="sxs-lookup"><span data-stu-id="12232-223">Do not test the returned object for nullity or falsity.</span></span> <span data-ttu-id="12232-224">Ele nunca é `null` , `false` ou `undefined` .</span><span class="sxs-lookup"><span data-stu-id="12232-224">It is never `null`, `false`, or `undefined`.</span></span>

<span data-ttu-id="12232-225">O exemplo de código a seguir tenta recuperar uma planilha do Excel chamada "data" usando o `getItemOrNullObject()` método.</span><span class="sxs-lookup"><span data-stu-id="12232-225">The following code sample attempts to retrieve an Excel worksheet named "Data" by using the `getItemOrNullObject()` method.</span></span> <span data-ttu-id="12232-226">Se uma planilha com esse nome não existir, será criada uma nova planilha.</span><span class="sxs-lookup"><span data-stu-id="12232-226">If a worksheet with that name does not exist, a new sheet is created.</span></span> <span data-ttu-id="12232-227">Observe que o código não carrega a `isNullObject` propriedade.</span><span class="sxs-lookup"><span data-stu-id="12232-227">Note that the code does not load the `isNullObject` property.</span></span> <span data-ttu-id="12232-228">O Office carrega automaticamente essa propriedade quando `context.sync` é chamado, portanto, você não precisa carregá-la explicitamente com algo como `datasheet.load('isNullObject')` .</span><span class="sxs-lookup"><span data-stu-id="12232-228">Office automatically loads this property when `context.sync` is called, so you do not need to explicitly load it with something like `datasheet.load('isNullObject')`.</span></span>

```js
var dataSheet = context.workbook.worksheets.getItemOrNullObject("Data");

return context.sync()
    .then(function () {
        if (dataSheet.isNullObject) {
            dataSheet = context.workbook.worksheets.add("Data");
        }

        // Set `dataSheet` to be the second worksheet in the workbook.
        dataSheet.position = 1;
    });
```

## <a name="see-also"></a><span data-ttu-id="12232-229">Confira também</span><span class="sxs-lookup"><span data-stu-id="12232-229">See also</span></span>

* [<span data-ttu-id="12232-230">Modelo de objeto comum de API JavaScript para Office</span><span class="sxs-lookup"><span data-stu-id="12232-230">Common JavaScript API object model</span></span>](office-javascript-api-object-model.md)
* [<span data-ttu-id="12232-231">Limites de recurso e otimização de desempenho para Suplementos do Office</span><span class="sxs-lookup"><span data-stu-id="12232-231">Resource limits and performance optimization for Office Add-ins</span></span>](../concepts/resource-limits-and-performance-optimization.md)
