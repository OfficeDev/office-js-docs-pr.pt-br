---
title: Conceitos fundamentais de programação com a API JavaScript do Excel
description: Use a API JavaScript do Excel para criar suplementos para o Excel.
ms.date: 07/13/2020
localization_priority: Priority
ms.openlocfilehash: 01e5fa1037719e89eed70f00e63431bbd445c213
ms.sourcegitcommit: 472b81642e9eb5fb2a55cd98a7b0826d37eb7f73
ms.translationtype: HT
ms.contentlocale: pt-BR
ms.lasthandoff: 07/17/2020
ms.locfileid: "45159413"
---
# <a name="fundamental-programming-concepts-with-the-excel-javascript-api"></a><span data-ttu-id="49975-103">Conceitos fundamentais de programação com a API JavaScript do Excel</span><span class="sxs-lookup"><span data-stu-id="49975-103">Fundamental programming concepts with the Excel JavaScript API</span></span>

<span data-ttu-id="49975-104">Este artigo descreve como usar a [API JavaScript do Excel](../reference/overview/excel-add-ins-reference-overview.md) para desenvolver suplementos para o Excel 2016 ou versões posteriores.</span><span class="sxs-lookup"><span data-stu-id="49975-104">This article describes how to use the [Excel JavaScript API](../reference/overview/excel-add-ins-reference-overview.md) to build add-ins for Excel 2016 or later.</span></span> <span data-ttu-id="49975-105">Ele apresenta os conceitos básicos que são fundamentais para usar a API e fornece orientações para executar tarefas específicas, como leitura ou gravação em um intervalo grande, atualização de todas as células do intervalo e muito mais.</span><span class="sxs-lookup"><span data-stu-id="49975-105">It introduces core concepts that are fundamental to using the API and provides guidance for performing specific tasks such as reading or writing to a large range, updating all cells in range, and more.</span></span>

## <a name="asynchronous-nature-of-excel-apis"></a><span data-ttu-id="49975-106">Natureza assíncrona das APIs do Excel</span><span class="sxs-lookup"><span data-stu-id="49975-106">Asynchronous nature of Excel APIs</span></span>

<span data-ttu-id="49975-p102">Os suplementos do Excel baseados na Web são executados dentro de um contêiner de navegador que é inserido no aplicativo do Office em plataformas baseadas em desktop, como Office no Windows e executado dentro de um iFrame HTML no Office na Web. Não é possível habilitar a API Office.js para interagir de modo síncrono com o host do Excel em todas as plataformas com suporte devido às considerações de desempenho. Desse modo, a chamada à API `sync()` no Office.js retorna uma [promessa](https://developer.mozilla.org/docs/Web/JavaScript/Reference/Global_Objects/Promise) que é resolvida quando o aplicativo Excel conclui as ações solicitadas de leitura ou gravação. Além disso, você pode enfileirar várias ações, como configurar propriedades ou invocar métodos, e executá-las como um lote de comandos com uma única chamada a `sync()`, em vez de enviar uma solicitação separada para cada ação. As seções a seguir descrevem como fazer isso usando as APIs `Excel.run()` e `sync()`.</span><span class="sxs-lookup"><span data-stu-id="49975-p102">The web-based Excel add-ins run inside a browser container that is embedded within the Office application on desktop-based platforms such as Office on Windows and runs inside an HTML iFrame in Office on the web. Enabling the Office.js API to interact synchronously with the Excel host across all supported platforms is not feasible due to performance considerations. Therefore, the `sync()` API call in Office.js returns a [promise](https://developer.mozilla.org/docs/Web/JavaScript/Reference/Global_Objects/Promise) that is resolved when the Excel application completes the requested read or write actions. Also, you can queue up multiple actions, such as setting properties or invoking methods, and run them as a batch of commands with a single call to `sync()`, rather than sending a separate request for each action. The following sections describe how to accomplish this using the `Excel.run()` and `sync()` APIs.</span></span>

## <a name="excelrun"></a><span data-ttu-id="49975-112">Excel.run</span><span class="sxs-lookup"><span data-stu-id="49975-112">Excel.run</span></span>

<span data-ttu-id="49975-p103">A `Excel.run` executa uma função em que você especifica as ações a serem executadas no modelo de objeto do Excel. A `Excel.run` cria automaticamente um contexto de solicitação que pode ser usado para sua interação com os objetos do Excel. Quando a `Excel.run` é concluída, uma promessa é resolvida e todos os objetos que foram alocados em tempo de execução são lançados automaticamente.</span><span class="sxs-lookup"><span data-stu-id="49975-p103">`Excel.run` executes a function where you specify the actions to perform against the Excel object model. `Excel.run` automatically creates a request context that you can use to interact with Excel objects. When `Excel.run` completes, a promise is resolved, and any objects that were allocated at runtime are automatically released.</span></span>

<span data-ttu-id="49975-p104">O exemplo a seguir mostra como usar o `Excel.run`. A instrução catch captura e registra erros que ocorrem dentro do `Excel.run`.</span><span class="sxs-lookup"><span data-stu-id="49975-p104">The following example shows how to use `Excel.run`. The catch statement catches and logs errors that occur within the `Excel.run`.</span></span>

```js
Excel.run(function (context) {
    // You can use the Excel JavaScript API here in the batch function
    // to execute actions on the Excel object model.
    console.log('Your code goes here.');
}).catch(function (error) {
    console.log('error: ' + error);
    if (error instanceof OfficeExtension.Error) {
        console.log('Debug info: ' + JSON.stringify(error.debugInfo));
    }
});
```

### <a name="run-options"></a><span data-ttu-id="49975-118">Executar opções</span><span class="sxs-lookup"><span data-stu-id="49975-118">Run options</span></span>

<span data-ttu-id="49975-119">`Excel.run` tem uma sobrecarga que recebe um objeto [RunOptions](/javascript/api/excel/excel.runoptions).</span><span class="sxs-lookup"><span data-stu-id="49975-119">`Excel.run` has an overload that takes in a [RunOptions](/javascript/api/excel/excel.runoptions) object.</span></span> <span data-ttu-id="49975-120">Este contém um conjunto de propriedades que afetam o comportamento de plataforma quando a função é executada.</span><span class="sxs-lookup"><span data-stu-id="49975-120">This contains a set of properties that affect platform behavior when the function runs.</span></span> <span data-ttu-id="49975-121">A propriedade a seguir tem suporte no momento:</span><span class="sxs-lookup"><span data-stu-id="49975-121">The following property is currently supported:</span></span>

- <span data-ttu-id="49975-122">`delayForCellEdit`: Determina se o Excel atrasa solicitação em lote até que o usuário sai do modo de edição de célula.</span><span class="sxs-lookup"><span data-stu-id="49975-122">`delayForCellEdit`: Determines whether Excel delays the batch request until the user exits cell edit mode.</span></span> <span data-ttu-id="49975-123">Quando **verdadeira**, a solicitação em lote é atrasada e executada quando o usuário sai do modo de edição de célula.</span><span class="sxs-lookup"><span data-stu-id="49975-123">When **true**, the batch request is delayed and runs when the user exits cell edit mode.</span></span> <span data-ttu-id="49975-124">Quando **falsa**, a solicitação em lote falha automaticamente se o usuário está no modo de edição de célula (causando um erro para alcançar o usuário).</span><span class="sxs-lookup"><span data-stu-id="49975-124">When **false**, the batch request automatically fails if the user is in cell edit mode (causing an error to reach the user).</span></span> <span data-ttu-id="49975-125">O comportamento padrão sem nenhuma propriedade `delayForCellEdit` especificada é equivalente a quando é **falsa**.</span><span class="sxs-lookup"><span data-stu-id="49975-125">The default behavior with no `delayForCellEdit` property specified is equivalent to when it is **false**.</span></span>

```js
Excel.run({ delayForCellEdit: true }, function (context) { ... })
```

## <a name="request-context"></a><span data-ttu-id="49975-126">Contexto de solicitação</span><span class="sxs-lookup"><span data-stu-id="49975-126">Request context</span></span>

<span data-ttu-id="49975-p107">O Excel e seu suplemento são executados em dois processos diferentes. Como eles usam diferentes ambientes de tempo de execução, os suplementos do Excel exigem um objeto `RequestContext` para conectar o suplemento aos objetos no Excel, como planilhas, intervalos, gráficos e tabelas.</span><span class="sxs-lookup"><span data-stu-id="49975-p107">Excel and your add-in run in two different processes. Since they use different runtime environments, Excel add-ins require a `RequestContext` object in order to connect your add-in to objects in Excel such as worksheets, ranges, charts, and tables.</span></span>

## <a name="proxy-objects"></a><span data-ttu-id="49975-129">Objetos proxy</span><span class="sxs-lookup"><span data-stu-id="49975-129">Proxy objects</span></span>

<span data-ttu-id="49975-p108">Os objetos JavaScript do Excel que você declara e usa em um suplemento são objetos proxy. Todos os métodos invocados, ou as propriedades definidas ou carregadas em objetos proxy são simplesmente adicionados a uma fila de comandos pendentes. Quando você chama o método `sync()` no contexto de solicitação (por exemplo, `context.sync()`), os comandos enfileirados são expedidos para o Excel e executados. A API JavaScript do Excel é basicamente centrada em lote. Você pode enfileirar quantas alterações desejar no contexto de solicitação e depois chamar o método `sync()` para executar o lote de comandos enfileirados.</span><span class="sxs-lookup"><span data-stu-id="49975-p108">The Excel JavaScript objects that you declare and use in an add-in are proxy objects. Any methods that you invoke or properties that you set or load on proxy objects are simply added to a queue of pending commands. When you call the `sync()` method on the request context (for example, `context.sync()`), the queued commands are dispatched to Excel and run. The Excel JavaScript API is fundamentally batch-centric. You can queue up as many changes as you wish on the request context, and then call the `sync()` method to run the batch of queued commands.</span></span>

<span data-ttu-id="49975-p109">Por exemplo, o trecho de código a seguir declara o objeto JavaScript local `selectedRange` para fazer referência a um intervalo selecionado no documento do Excel e, em seguida, define algumas propriedades nesse objeto. O objeto `selectedRange` é um objeto proxy, de modo que as propriedades que são definidas e o método que é invocado nesse objeto não serão refletidos no documento do Excel até que o suplemento chame `context.sync()`.</span><span class="sxs-lookup"><span data-stu-id="49975-p109">For example, the following code snippet declares the local JavaScript object `selectedRange` to reference a selected range in the Excel document, and then sets some properties on that object. The `selectedRange` object is a proxy object, so the properties that are set and method that is invoked on that object will not be reflected in the Excel document until your add-in calls `context.sync()`.</span></span>

```js
var selectedRange = context.workbook.getSelectedRange();
selectedRange.format.fill.color = "#4472C4";
selectedRange.format.font.color = "white";
selectedRange.format.autofitColumns();
```

### <a name="sync"></a><span data-ttu-id="49975-137">sync()</span><span class="sxs-lookup"><span data-stu-id="49975-137">sync()</span></span>

<span data-ttu-id="49975-p110">Chamar o método `sync()` no contexto de solicitação sincroniza o estado entre objetos proxy e objetos no documento do Excel. O método `sync()` executa todos os comandos que são enfileirados no contexto de solicitação e recupera valores para qualquer propriedade que deva ser carregada nos objetos proxy. O método `sync()` é executado de modo assíncrono e retorna uma [promessa](https://developer.mozilla.org/docs/Web/JavaScript/Reference/Global_Objects/Promise), que é resolvida quando o método `sync()` é concluído.</span><span class="sxs-lookup"><span data-stu-id="49975-p110">Calling the `sync()` method on the request context synchronizes the state between proxy objects and objects in the Excel document. The `sync()` method runs any commands that are queued on the request context and retrieves values for any properties that should be loaded on the proxy objects. The `sync()` method executes asynchronously and returns a [promise](https://developer.mozilla.org/docs/Web/JavaScript/Reference/Global_Objects/Promise), which is resolved when the `sync()` method completes.</span></span>

<span data-ttu-id="49975-141">O exemplo a seguir mostra uma função de lote que define um objeto proxy JavaScript local (`selectedRange`), carrega uma propriedade desse objeto e, em seguida, usa o padrão Promessas do JavaScript para chamar `context.sync()` a fim de sincronizar o estado entre objetos proxy e objetos no documento do Excel.</span><span class="sxs-lookup"><span data-stu-id="49975-141">The following example shows a batch function that defines a local JavaScript proxy object (`selectedRange`), loads a property of that object, and then uses the JavaScript Promises pattern to call `context.sync()` to synchronize the state between proxy objects and objects in the Excel document.</span></span>

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

<span data-ttu-id="49975-142">No exemplo anterior, `selectedRange` é definido e sua propriedade `address` é carregada quando `context.sync()` é chamado.</span><span class="sxs-lookup"><span data-stu-id="49975-142">In the previous example, `selectedRange` is set and its `address` property is loaded when `context.sync()` is called.</span></span>

<span data-ttu-id="49975-143">Como `sync()` é uma operação assíncrona que retorna uma promessa, você deve sempre `return` a promessa (em JavaScript).</span><span class="sxs-lookup"><span data-stu-id="49975-143">Because `sync()` is an asynchronous operation that returns a promise, you should always `return` the promise (in JavaScript).</span></span> <span data-ttu-id="49975-144">Isso garante que a operação `sync()` seja concluída antes que o script continue em execução.</span><span class="sxs-lookup"><span data-stu-id="49975-144">Doing so ensures that the `sync()` operation completes before the script continues to run.</span></span> <span data-ttu-id="49975-145">Para obter mais informações sobre como otimizar o desempenho com o `sync()`, consulte [Otimização de desempenho da API JavaScript do Excel](../excel/performance.md).</span><span class="sxs-lookup"><span data-stu-id="49975-145">For more information about optimizing performance with `sync()`, see [Excel JavaScript API performance optimization](../excel/performance.md).</span></span>

### <a name="load"></a><span data-ttu-id="49975-146">load()</span><span class="sxs-lookup"><span data-stu-id="49975-146">load()</span></span>

<span data-ttu-id="49975-p112">Para que você possa ler as propriedades de um objeto proxy, é preciso carregar explicitamente as propriedades para popular o objeto proxy com dados do documento do Excel e chamar `context.sync()`. Por exemplo, se você criar um objeto proxy para fazer referência a um intervalo selecionado e, em seguida, quiser ler a propriedade `address` do intervalo selecionado, será preciso carregar a propriedade `address` para que seja possível lê-la. Para solicitar que as propriedades de um objeto proxy sejam carregadas, chame o método `load()` no objeto e especifique as propriedades a serem carregadas.</span><span class="sxs-lookup"><span data-stu-id="49975-p112">Before you can read the properties of a proxy object, you must explicitly load the properties to populate the proxy object with data from the Excel document, and then call `context.sync()`. For example, if you create a proxy object to reference a selected range, and then want to read the selected range's `address` property, you need to load the `address` property before you can read it. To request properties of a proxy object be loaded, call the `load()` method on the object and specify the properties to load.</span></span>

> [!NOTE]
> <span data-ttu-id="49975-p113">Se estiver apenas chamando métodos ou definindo propriedades em um objeto proxy, você não precisa chamar o método `load()`. O método `load()` só é necessário quando você deseja ler propriedades em um objeto proxy.</span><span class="sxs-lookup"><span data-stu-id="49975-p113">If you are only calling methods or setting properties on a proxy object, you do not need to call the `load()` method. The `load()` method is only required when you want to read properties on a proxy object.</span></span>

<span data-ttu-id="49975-p114">Assim como as solicitações para definir propriedades ou invocar métodos em objetos proxy, as solicitações para carregar propriedades em objetos proxy são adicionadas à fila de comandos pendentes no contexto de solicitação, sendo executadas na próxima vez que você chamar o método `sync()`. É possível enfileirar quantas chamadas de `load()` forem necessárias no contexto de solicitação.</span><span class="sxs-lookup"><span data-stu-id="49975-p114">Just like requests to set properties or invoke methods on proxy objects, requests to load properties on proxy objects get added to the queue of pending commands on the request context, which will run the next time you call the `sync()` method. You can queue up as many `load()` calls on the request context as necessary.</span></span>

<span data-ttu-id="49975-154">No exemplo a seguir, somente propriedades específicas do intervalo são carregadas.</span><span class="sxs-lookup"><span data-stu-id="49975-154">In the following example, only specific properties of the range are loaded.</span></span>

```js
Excel.run(function (context) {
    var sheetName = 'Sheet1';
    var rangeAddress = 'A1:B2';
    var myRange = context.workbook.worksheets.getItem(sheetName).getRange(rangeAddress);

    myRange.load(['address', 'format/*', 'format/fill', 'entireRow' ]);

    return context.sync()
      .then(function () {
        console.log (myRange.address);              // ok
        console.log (myRange.format.wrapText);      // ok
        console.log (myRange.format.fill.color);    // ok
        //console.log (myRange.format.font.color);  // not ok as it was not loaded
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

<span data-ttu-id="49975-155">No exemplo anterior, como `format/font` não está especificado na chamada para `myRange.load()`, a propriedade `format.font.color` não pode ser lida.</span><span class="sxs-lookup"><span data-stu-id="49975-155">In the previous example, because `format/font` is not specified in the call to `myRange.load()`, the `format.font.color` property cannot be read.</span></span>

<span data-ttu-id="49975-156">Para otimizar o desempenho, você deve especificar explicitamente as propriedades e as relações a serem carregadas ao usar o método `load()` em um objeto, como abrangido em [Otimizações do desempenho da API JavaScript do Excel](performance.md).</span><span class="sxs-lookup"><span data-stu-id="49975-156">To optimize performance, you should explicitly specify the properties to load when using the `load()` method on an object, as covered in [Excel JavaScript API performance optimizations](performance.md).</span></span> <span data-ttu-id="49975-157">Para obter mais informações sobre o método `load()`, consulte [Conceitos avançados de programação com a API JavaScript do Excel](excel-add-ins-advanced-concepts.md).</span><span class="sxs-lookup"><span data-stu-id="49975-157">For more information about the `load()` method, see [Advanced programming concepts with the Excel JavaScript API](excel-add-ins-advanced-concepts.md).</span></span>

## <a name="null-or-blank-property-values"></a><span data-ttu-id="49975-158">Valores de propriedade nulos ou em branco</span><span class="sxs-lookup"><span data-stu-id="49975-158">null or blank property values</span></span>

### <a name="null-input-in-2-d-array"></a><span data-ttu-id="49975-159">entrada nula em uma matriz 2D</span><span class="sxs-lookup"><span data-stu-id="49975-159">null input in 2-D Array</span></span>

<span data-ttu-id="49975-p116">No Excel, um intervalo é representado por uma matriz 2D, onde a primeira dimensão é linhas e a segunda dimensão é colunas. Para definir valores, o formato do número ou a fórmula apenas para células específicas em um intervalo, especifique os valores, o formato do número ou a fórmula para essas células na matriz 2D, bem como `null` para todas as outras células na matriz 2D.</span><span class="sxs-lookup"><span data-stu-id="49975-p116">In Excel, a range is represented by a 2-D array, where the first dimension is rows and the second dimension is columns. To set values, number format, or formula for only specific cells within a range, specify the values, number format, or formula for those cells in the 2-D array, and specify `null` for all other cells in the 2-D array.</span></span>

<span data-ttu-id="49975-p117">Por exemplo, para atualizar o formato do número apenas para uma célula em um intervalo e manter o formato de número existente para todas as outras células no intervalo, especifique o novo formato de número para a célula a ser atualizada e `null` para todas as outras células. O trecho de código a seguir define um novo formato de número para a quarta célula no intervalo e não altera o formato de número para as primeiras três células no intervalo.</span><span class="sxs-lookup"><span data-stu-id="49975-p117">For example, to update the number format for only one cell within a range, and retain the existing number format for all other cells in the range, specify the new number format for the cell to update, and specify `null` for all other cells. The following code snippet sets a new number format for the fourth cell in the range, and leaves the number format unchanged for the first three cells in the range.</span></span>

```js
range.values = [['Eurasia', '29.96', '0.25', '15-Feb' ]];
range.numberFormat = [[null, null, null, 'm/d/yyyy;@']];
```

### <a name="null-input-for-a-property"></a><span data-ttu-id="49975-164">entrada nula para uma propriedade</span><span class="sxs-lookup"><span data-stu-id="49975-164">null input for a property</span></span>

<span data-ttu-id="49975-p118">`null` não é uma entrada válida para uma propriedade única. Por exemplo, o trecho de código a seguir não é válido, pois a propriedade `values` do intervalo não pode ser definida como `null`.</span><span class="sxs-lookup"><span data-stu-id="49975-p118">`null` is not a valid input for single property. For example, the following code snippet is not valid, as the `values` property of the range cannot be set to `null`.</span></span>

```js
range.values = null;
```

<span data-ttu-id="49975-167">Da mesma forma, o seguinte snippet de código não é válido, pois `null` não é um valor válido para a propriedade `color`.</span><span class="sxs-lookup"><span data-stu-id="49975-167">Likewise, the following code snippet is not valid, as `null` is not a valid value for the `color` property.</span></span>

```js
range.format.fill.color =  null;
```

### <a name="null-property-values-in-the-response"></a><span data-ttu-id="49975-168">Valores da propriedade nula na resposta</span><span class="sxs-lookup"><span data-stu-id="49975-168">null property values in the response</span></span>

<span data-ttu-id="49975-p119">A formatação de propriedades como `size` e `color` conterá valores `null` na resposta quando valores diferentes existirem no intervalo especificado. Por exemplo, se você recuperar um intervalo e carregar sua propriedade `format.font.color`:</span><span class="sxs-lookup"><span data-stu-id="49975-p119">Formatting properties such as `size` and `color` will contain `null` values in the response when different values exist in the specified range. For example, if you retrieve a range and load its `format.font.color` property:</span></span>

- <span data-ttu-id="49975-171">Se todas as células no intervalo tiverem a mesma cor de fonte, `range.format.font.color` especificará essa cor.</span><span class="sxs-lookup"><span data-stu-id="49975-171">If all cells in the range have the same font color, `range.format.font.color` specifies that color.</span></span>
- <span data-ttu-id="49975-172">Se houver várias cores de fonte dentro do intervalo, `range.format.font.color` será `null`.</span><span class="sxs-lookup"><span data-stu-id="49975-172">If multiple font colors are present within the range, `range.format.font.color` is `null`.</span></span>

### <a name="blank-input-for-a-property"></a><span data-ttu-id="49975-173">Entrada em branco para uma propriedade</span><span class="sxs-lookup"><span data-stu-id="49975-173">Blank input for a property</span></span>

<span data-ttu-id="49975-p120">Quando você especificar um valor em branco para uma propriedade (isto é, duas aspas sem espaço entre elas `''`), ele será interpretado como uma instrução para limpar ou redefinir a propriedade. Por exemplo:</span><span class="sxs-lookup"><span data-stu-id="49975-p120">When you specify a blank value for a property (i.e., two quotation marks with no space in-between `''`), it will be interpreted as an instruction to clear or reset the property. For example:</span></span>

- <span data-ttu-id="49975-176">Se você especificar um valor em branco para a propriedade `values` de um intervalo, o conteúdo do intervalo será apagado.</span><span class="sxs-lookup"><span data-stu-id="49975-176">If you specify a blank value for the `values` property of a range, the content of the range is cleared.</span></span>

- <span data-ttu-id="49975-177">Se você especificar um valor em branco para a propriedade `numberFormat`, o formato de número será redefinido para `General`.</span><span class="sxs-lookup"><span data-stu-id="49975-177">If you specify a blank value for the `numberFormat` property, the number format is reset to `General`.</span></span>

- <span data-ttu-id="49975-178">Se você especificar um valor em branco para a propriedade `formula` e a propriedade `formulaLocale`, os valores de fórmula serão apagados.</span><span class="sxs-lookup"><span data-stu-id="49975-178">If you specify a blank value for the `formula` property and `formulaLocale` property, the formula values are cleared.</span></span>

### <a name="blank-property-values-in-the-response"></a><span data-ttu-id="49975-179">Valores da propriedade em branco na resposta</span><span class="sxs-lookup"><span data-stu-id="49975-179">Blank property values in the response</span></span>

<span data-ttu-id="49975-p121">Para operações de leitura, um valor de propriedade em branco na resposta (isto é, duas aspas sem espaço entre elas `''`) indica que a célula não contém dados nem valor. No primeiro exemplo abaixo, a primeira e a última célula no intervalo não contêm dados. No segundo exemplo, as primeiras duas células no intervalo não contêm uma fórmula.</span><span class="sxs-lookup"><span data-stu-id="49975-p121">For read operations, a blank property value in the response (i.e., two quotation marks with no space in-between `''`) indicates that cell contains no data or value. In the first example below, the first and last cell in the range contain no data. In the second example, the first two cells in the range do not contain a formula.</span></span>

```js
range.values = [['', 'some', 'data', 'in', 'other', 'cells', '']];
```

```js
range.formula = [['', '', '=Rand()']];
```

## <a name="read-or-write-to-an-unbounded-range"></a><span data-ttu-id="49975-183">Ler ou gravar em um intervalo não limitado</span><span class="sxs-lookup"><span data-stu-id="49975-183">Read or write to an unbounded range</span></span>

### <a name="read-an-unbounded-range"></a><span data-ttu-id="49975-184">Ler um intervalo não limitado</span><span class="sxs-lookup"><span data-stu-id="49975-184">Read an unbounded range</span></span>

<span data-ttu-id="49975-p122">Um endereço de intervalo não limitado é um endereço de intervalo que especifica colunas ou linhas inteiras. Por exemplo:</span><span class="sxs-lookup"><span data-stu-id="49975-p122">An unbounded range address is a range address that specifies either entire column(s) or entire row(s). For example:</span></span>

- <span data-ttu-id="49975-187">Endereços de intervalo composto por colunas inteiras:</span><span class="sxs-lookup"><span data-stu-id="49975-187">Range addresses comprised of entire column(s):</span></span><ul><li>`C:C`</li><li>`A:F`</li></ul>
- <span data-ttu-id="49975-188">Endereços de intervalo composto por linhas inteiras:</span><span class="sxs-lookup"><span data-stu-id="49975-188">Range addresses comprised of entire row(s):</span></span><ul><li>`2:2`</li><li>`1:4`</li></ul>

<span data-ttu-id="49975-p123">Quando uma API faz uma solicitação para recuperar um intervalo não limitado (por exemplo, `getRange('C:C')`), a resposta conterá valores `null` para as propriedades no nível de célula, como `values`, `text`, `numberFormat` e `formula`. Outras propriedades do intervalo, como `address` e `cellCount`, conterão valores válidos para o intervalo não limitado.</span><span class="sxs-lookup"><span data-stu-id="49975-p123">When the API makes a request to retrieve an unbounded range (for example, `getRange('C:C')`), the response will contain `null` values for cell-level properties such as `values`, `text`, `numberFormat`, and `formula`. Other properties of the range, such as `address` and `cellCount`, will contain valid values for the unbounded range.</span></span>

### <a name="write-to-an-unbounded-range"></a><span data-ttu-id="49975-191">Gravar em um intervalo não limitado</span><span class="sxs-lookup"><span data-stu-id="49975-191">Write to an unbounded range</span></span>

<span data-ttu-id="49975-p124">Não é possível definir propriedades no nível de célula, como `values`, `numberFormat` e `formula`, no intervalo não limitado, pois a solicitação de entrada é muito grande. Por exemplo, o trecho de código a seguir não é válida porque ele tenta especificar `values` para um intervalo não limitado. A API retornará um erro se você tentar definir as propriedades no nível de célula para um intervalo não limitado.</span><span class="sxs-lookup"><span data-stu-id="49975-p124">You cannot set cell-level properties such as `values`, `numberFormat`, and `formula` on unbounded range because the input request is too large. For example, the following code snippet is not valid because it attempts to specify `values` for an unbounded range. The API will return an error if you attempt to set cell-level properties for an unbounded range.</span></span>

```js
var range = context.workbook.worksheets.getActiveWorksheet().getRange('A:B');
range.values = 'Due Date';
```

## <a name="read-or-write-to-a-large-range"></a><span data-ttu-id="49975-195">Ler ou gravar em um intervalo grande</span><span class="sxs-lookup"><span data-stu-id="49975-195">Read or write to a large range</span></span>

<span data-ttu-id="49975-p125">Se um intervalo contiver um grande número de células, valores, formatos de número e/ou fórmulas, talvez não seja possível executar operações de API nesse intervalo. A API sempre fará a melhor tentativa de executar a operação solicitada em um intervalo (isto é, para recuperar ou gravar os dados especificados), mas tentar executar operações de leitura ou gravação para um intervalo grande pode resultar em um erro de API devido à utilização excessiva de recursos. Para evitar tais erros, é recomendável executar operações de leitura ou gravação separadas para subconjuntos menores de um intervalo grande, em vez de tentar executar uma única operação de leitura ou gravação em um intervalo grande.</span><span class="sxs-lookup"><span data-stu-id="49975-p125">If a range contains a large number of cells, values, number formats, and/or formulas, it may not be possible to run API operations on that range. The API will always make a best attempt to run the requested operation on a range (i.e., to retrieve or write the specified data), but attempting to perform read or write operations for a large range may result in an API error due to excessive resource utilization. To avoid such errors, we recommend that you run separate read or write operations for smaller subsets of a large range, instead of attempting to run a single read or write operation on a large range.</span></span>

<span data-ttu-id="49975-199">Para detalhes sobre as limitações do sistema, consulte [Limites de transferência de dados do Excel](../develop/common-coding-issues.md#excel-data-transfer-limits).</span><span class="sxs-lookup"><span data-stu-id="49975-199">For details on the system limitations, see [Excel data transfer limits](../develop/common-coding-issues.md#excel-data-transfer-limits).</span></span>

## <a name="handle-errors"></a><span data-ttu-id="49975-200">Lidar com erros</span><span class="sxs-lookup"><span data-stu-id="49975-200">Handle errors</span></span>

<span data-ttu-id="49975-201">Quando ocorre um erro de API, a API retorna um objeto `error` que contém um código e uma mensagem.</span><span class="sxs-lookup"><span data-stu-id="49975-201">When an API error occurs, the API returns an `error` object that contains a code and a message.</span></span> <span data-ttu-id="49975-202">Para saber mais sobre o tratamento de erros, incluindo uma lista de erros da API, confira [Tratamento de erro](excel-add-ins-error-handling.md).</span><span class="sxs-lookup"><span data-stu-id="49975-202">For detailed information about error handling, including a list of API errors, see [Error handling](excel-add-ins-error-handling.md).</span></span>

## <a name="see-also"></a><span data-ttu-id="49975-203">Confira também</span><span class="sxs-lookup"><span data-stu-id="49975-203">See also</span></span>

- [<span data-ttu-id="49975-204">Crie seu primeiro suplemento do Excel</span><span class="sxs-lookup"><span data-stu-id="49975-204">Build your first Excel add-in</span></span>](../quickstarts/excel-quickstart-jquery.md)
- [<span data-ttu-id="49975-205">Exemplos de código de suplementos do Excel</span><span class="sxs-lookup"><span data-stu-id="49975-205">Excel add-ins code samples</span></span>](https://developer.microsoft.com/office/gallery/?filterBy=Samples,Excel)
- [<span data-ttu-id="49975-206">Conceitos fundamentais de programação com a API JavaScript do Excel</span><span class="sxs-lookup"><span data-stu-id="49975-206">Advanced programming concepts with the Excel JavaScript API</span></span>](excel-add-ins-advanced-concepts.md)
- [<span data-ttu-id="49975-207">Otimização de desempenho do da API JavaScript do Excel</span><span class="sxs-lookup"><span data-stu-id="49975-207">Excel JavaScript API performance optimization</span></span>](../excel/performance.md)
- [<span data-ttu-id="49975-208">Referência da API JavaScript do Excel</span><span class="sxs-lookup"><span data-stu-id="49975-208">Excel JavaScript API reference</span></span>](../reference/overview/excel-add-ins-reference-overview.md)
- <span data-ttu-id="49975-209">[Problemas comuns de codificação e comportamentos inesperados da plataforma](../develop/common-coding-issues.md).</span><span class="sxs-lookup"><span data-stu-id="49975-209">[Common coding issues and unexpected platform behaviors](../develop/common-coding-issues.md).</span></span>
