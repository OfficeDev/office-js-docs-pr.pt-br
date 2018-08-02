---
title: Principais conceitos da API JavaScript do Excel
description: ''
ms.date: 12/04/2017
ms.openlocfilehash: fb22ae41718c459366a628c8f06531cc6978a178
ms.sourcegitcommit: bc68b4cf811b45e8b8d1cbd7c8d2867359ab671b
ms.translationtype: HT
ms.contentlocale: pt-BR
ms.lasthandoff: 08/02/2018
ms.locfileid: "21703837"
---
# <a name="excel-javascript-api-core-concepts"></a><span data-ttu-id="1df3b-102">Principais conceitos da API JavaScript do Excel</span><span class="sxs-lookup"><span data-stu-id="1df3b-102">Excel JavaScript API core concepts</span></span>
 
<span data-ttu-id="1df3b-103">Este artigo descreve como usar a [API JavaScript do Excel](https://dev.office.com/reference/add-ins/excel/excel-add-ins-reference-overview) para desenvolver suplementos para o Excel 2016.</span><span class="sxs-lookup"><span data-stu-id="1df3b-103">This article describes how to use the [Excel JavaScript API](https://dev.office.com/reference/add-ins/excel/excel-add-ins-reference-overview) to build add-ins for Excel 2016.</span></span> <span data-ttu-id="1df3b-104">Ele apresenta os conceitos básicos que são fundamentais para usar a API e fornece orientações para executar tarefas específicas, como leitura ou gravação em um intervalo grande, atualização de todas as células do intervalo e muito mais.</span><span class="sxs-lookup"><span data-stu-id="1df3b-104">It introduces core concepts that are fundamental to using the API and provides guidance for performing specific tasks such as reading or writing to a large range, updating all cells in range, and more.</span></span>

## <a name="asynchronous-nature-of-excel-apis"></a><span data-ttu-id="1df3b-105">Natureza assíncrona das APIs do Excel</span><span class="sxs-lookup"><span data-stu-id="1df3b-105">Asynchronous nature of Excel APIs</span></span>

<span data-ttu-id="1df3b-106">Os suplementos do Excel baseados na Web são executados dentro de um contêiner de navegador que é inserido no aplicativo do Office em plataformas baseadas em desktop, como Office para Windows, e executado dentro de um iFrame HTML no Office Online.</span><span class="sxs-lookup"><span data-stu-id="1df3b-106">The web-based Excel add-ins run inside a browser container that is embedded within the Office application on desktop-based platforms such as Office for Windows and runs inside an HTML iFrame in Office Online.</span></span> <span data-ttu-id="1df3b-107">Não é possível habilitar a API Office.js para interagir de modo síncrono com o host do Excel em todas as plataformas suportadas devido às considerações de desempenho.</span><span class="sxs-lookup"><span data-stu-id="1df3b-107">Enabling the Office.js API to interact synchronously with the Excel host across all supported platforms is not feasible due to performance considerations.</span></span> <span data-ttu-id="1df3b-108">Desse modo, a chamada à API **sync()** na Office.js retorna uma [promessa](https://developer.mozilla.org/en-US/docs/Web/JavaScript/Reference/Global_Objects/Promise) que é resolvida quando o aplicativo Excel conclui as ações solicitadas de leitura ou gravação.</span><span class="sxs-lookup"><span data-stu-id="1df3b-108">Therefore, the **sync()** API call in Office.js returns a [promise](https://developer.mozilla.org/en-US/docs/Web/JavaScript/Reference/Global_Objects/Promise) that is resolved when the Excel application completes the requested read or write actions.</span></span> <span data-ttu-id="1df3b-109">Além disso, você pode enfileirar várias ações, como configurar propriedades ou invocar métodos, e executá-las como um lote de comandos com uma única chamada a **sync()**, em vez de enviar uma solicitação separada para cada ação.</span><span class="sxs-lookup"><span data-stu-id="1df3b-109">Also, you can queue up multiple actions, such as setting properties or invoking methods, and run them as a batch of commands with a single call to **sync()**, rather than sending a separate request for each action.</span></span> <span data-ttu-id="1df3b-110">As seções a seguir descrevem como fazer isso usando as APIs **Excel.run()** e **sync()**.</span><span class="sxs-lookup"><span data-stu-id="1df3b-110">The following sections describe how to accomplish this using the **Excel.run()** and **sync()** APIs.</span></span>
 
## <a name="excelrun"></a><span data-ttu-id="1df3b-111">Excel.run</span><span class="sxs-lookup"><span data-stu-id="1df3b-111">Excel.run</span></span>
 
<span data-ttu-id="1df3b-112">A **Excel.run** executa uma função em que você especifica as ações a serem executadas no modelo de objeto do Excel.</span><span class="sxs-lookup"><span data-stu-id="1df3b-112">**Excel.run** executes a function where you specify the actions to perform against the Excel object model.</span></span> <span data-ttu-id="1df3b-113">A **Excel.run** cria automaticamente um contexto de solicitação que pode ser usado para sua interação com os objetos do Excel.</span><span class="sxs-lookup"><span data-stu-id="1df3b-113">**Excel.run** automatically creates a request context that you can use to interact with Excel objects.</span></span> <span data-ttu-id="1df3b-114">Quando a **Excel.run** é concluída, uma promessa é resolvida e todos os objetos que foram alocados em tempo de execução são lançados automaticamente.</span><span class="sxs-lookup"><span data-stu-id="1df3b-114">When **Excel.run** completes, a promise is resolved, and any objects that were allocated at runtime are automatically released.</span></span>
 
<span data-ttu-id="1df3b-115">O exemplo a seguir mostra como usar a **Excel.run**.</span><span class="sxs-lookup"><span data-stu-id="1df3b-115">The following example shows how to use **Excel.run**.</span></span> <span data-ttu-id="1df3b-116">A instrução catch captura e grava em log os erros que ocorrem na **Excel.run**.</span><span class="sxs-lookup"><span data-stu-id="1df3b-116">The catch statement catches and logs errors that occur within the **Excel.run**.</span></span>
 
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

## <a name="request-context"></a><span data-ttu-id="1df3b-117">Contexto de solicitação</span><span class="sxs-lookup"><span data-stu-id="1df3b-117">Request context</span></span>
 
<span data-ttu-id="1df3b-p105">O Excel e seu suplemento são executados em dois processos diferentes. Como eles usam diferentes ambientes de tempo de execução, os suplementos do Excel exigem um objeto **RequestContext** para conectar o suplemento aos objetos no Excel, como planilhas, intervalos, gráficos e tabelas.</span><span class="sxs-lookup"><span data-stu-id="1df3b-p105">Excel and your add-in run in two different processes. Since they use different runtime environments, Excel add-ins require a **RequestContext** object in order to connect your add-in to objects in Excel such as worksheets, ranges, charts, and tables.</span></span>
 
## <a name="proxy-objects"></a><span data-ttu-id="1df3b-120">Objetos proxy</span><span class="sxs-lookup"><span data-stu-id="1df3b-120">Proxy objects</span></span>
 
<span data-ttu-id="1df3b-121">Os objetos JavaScript do Excel que você declara e usa em um suplemento são objetos proxy.</span><span class="sxs-lookup"><span data-stu-id="1df3b-121">The Excel JavaScript objects that you declare and use in an add-in are proxy objects.</span></span> <span data-ttu-id="1df3b-122">Todos os métodos invocados, ou as propriedades definidas ou carregadas em objetos proxy são simplesmente adicionados a uma fila de comandos pendentes.</span><span class="sxs-lookup"><span data-stu-id="1df3b-122">Any methods that you invoke or properties that you set or load on proxy objects are simply added to a queue of pending commands.</span></span> <span data-ttu-id="1df3b-123">Quando você chama o método **sync()** no contexto de solicitação (por exemplo, `context.sync()`), os comandos enfileirados são expedidos para o Excel e executados.</span><span class="sxs-lookup"><span data-stu-id="1df3b-123">When you call the **sync()** method on the request context (for example, `context.sync()`), the queued commands are dispatched to Excel and run.</span></span> <span data-ttu-id="1df3b-124">A API JavaScript do Excel é basicamente centrada em lote.</span><span class="sxs-lookup"><span data-stu-id="1df3b-124">The Excel JavaScript API is fundamentally batch-centric.</span></span> <span data-ttu-id="1df3b-125">Você pode enfileirar quantas alterações desejar no contexto de solicitação e depois chamar o método **sync()** para executar o lote de comandos enfileirados.</span><span class="sxs-lookup"><span data-stu-id="1df3b-125">You can queue up as many changes as you wish on the request context, and then call the **sync()** method to run the batch of queued commands.</span></span>
 
<span data-ttu-id="1df3b-126">Por exemplo, o trecho de código a seguir declara o objeto JavaScript local **selectedRange** para fazer referência a um intervalo selecionado no documento do Excel e, em seguida, define algumas propriedades nesse objeto.</span><span class="sxs-lookup"><span data-stu-id="1df3b-126">For example, the following code snippet declares the local JavaScript object **selectedRange** to reference a selected range in the Excel document, and then sets some properties on that object.</span></span> <span data-ttu-id="1df3b-127">O objeto **selectedRange** é um objeto proxy, de modo que as propriedades que são definidas e o método que é invocado nesse objeto não serão refletidos no documento do Excel até que o suplemento chame **context.sync()**.</span><span class="sxs-lookup"><span data-stu-id="1df3b-127">The **selectedRange** object is a proxy object, so the properties that are set and method that is invoked on that object will not be reflected in the Excel document until your add-in calls **context.sync()**.</span></span>
 
```js
const selectedRange = context.workbook.getSelectedRange();
selectedRange.format.fill.color = "#4472C4";
selectedRange.format.font.color = "white";
selectedRange.format.autofitColumns();
```
 
### <a name="sync"></a><span data-ttu-id="1df3b-128">sync()</span><span class="sxs-lookup"><span data-stu-id="1df3b-128">sync()</span></span>
 
<span data-ttu-id="1df3b-129">Chamar o método **sync()** no contexto de solicitação sincroniza o estado entre objetos proxy e objetos no documento do Excel.</span><span class="sxs-lookup"><span data-stu-id="1df3b-129">Calling the **sync()** method on the request context synchronizes the state between proxy objects and objects in the Excel document.</span></span> <span data-ttu-id="1df3b-130">O método **sync()** executa todos os comandos que são enfileirados no contexto de solicitação e recupera valores para qualquer propriedade que deva ser carregada nos objetos proxy.</span><span class="sxs-lookup"><span data-stu-id="1df3b-130">The **sync()** method runs any commands that are queued on the request context and retrieves values for any properties that should be loaded on the proxy objects.</span></span> <span data-ttu-id="1df3b-131">O método **sync()** é executado de modo assíncrono e retorna uma [promessa](https://developer.mozilla.org/en-US/docs/Web/JavaScript/Reference/Global_Objects/Promise), que é resolvida quando o método **sync()** é concluído.</span><span class="sxs-lookup"><span data-stu-id="1df3b-131">The **sync()** method executes asynchronously and returns a [promise](https://developer.mozilla.org/en-US/docs/Web/JavaScript/Reference/Global_Objects/Promise), which is resolved when the **sync()** method completes.</span></span>
 
<span data-ttu-id="1df3b-132">O exemplo a seguir mostra uma função de lote que define um objeto proxy JavaScript local (**selectedRange**), carrega uma propriedade desse objeto e, em seguida, usa o padrão Promessas do JavaScript para chamar **context.sync()** a fim de sincronizar o estado entre objetos proxy e objetos no documento do Excel.</span><span class="sxs-lookup"><span data-stu-id="1df3b-132">The following example shows a batch function that defines a local JavaScript proxy object (**selectedRange**), loads a property of that object, and then uses the JavaScript Promises pattern to call **context.sync()** to synchronize the state between proxy objects and objects in the Excel document.</span></span>
 
```js
Excel.run(function (context) {
  const selectedRange = context.workbook.getSelectedRange();
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
 
<span data-ttu-id="1df3b-133">No exemplo anterior, **selectedRange** está definido e sua propriedade **address** é carregada quando **context.sync()** é chamado.</span><span class="sxs-lookup"><span data-stu-id="1df3b-133">In the previous example, **selectedRange** is set and its **address** property is loaded when **context.sync()** is called.</span></span>
 
<span data-ttu-id="1df3b-134">Como **sync()** é uma operação assíncrona que retorna uma promessa, você sempre deve **retornar** a promessa (no JavaScript).</span><span class="sxs-lookup"><span data-stu-id="1df3b-134">Because **sync()** is an asynchronous operation that returns a promise, you should always **return** the promise (in JavaScript).</span></span> <span data-ttu-id="1df3b-135">Isso garante que a operação **sync()** seja concluída antes que o script continue sendo executado.</span><span class="sxs-lookup"><span data-stu-id="1df3b-135">Doing so ensures that the **sync()** operation completes before the script continues to run.</span></span> <span data-ttu-id="1df3b-136">Para obter mais informações sobre como otimizar o desempenho com **sync()**, confira [Otimização de desempenho da API JavaScript do Excel](https://docs.microsoft.com/office/dev/add-ins/excel/performance).</span><span class="sxs-lookup"><span data-stu-id="1df3b-136">For more information about optimizing performance with **sync()**, see [Excel JavaScript API performance optimization](https://docs.microsoft.com/office/dev/add-ins/excel/performance).</span></span>
 
### <a name="load"></a><span data-ttu-id="1df3b-137">load()</span><span class="sxs-lookup"><span data-stu-id="1df3b-137">load()</span></span>
 
<span data-ttu-id="1df3b-138">Para que você possa ler as propriedades de um objeto proxy, é preciso carregar explicitamente as propriedades para popular o objeto proxy com dados do documento do Excel e chamar **context.sync()**.</span><span class="sxs-lookup"><span data-stu-id="1df3b-138">Before you can read the properties of a proxy object, you must explicitly load the properties to populate the proxy object with data from the Excel document, and then call **context.sync()**.</span></span> <span data-ttu-id="1df3b-139">Por exemplo, se você criar um objeto proxy para fazer referência a um intervalo selecionado e, em seguida, quiser ler a propriedade **address** do intervalo selecionado, será preciso carregar a propriedade **address** para que seja possível lê-la.</span><span class="sxs-lookup"><span data-stu-id="1df3b-139">For example, if you create a proxy object to reference a selected range, and then want to read the selected range's **address** property, you need to load the **address** property before you can read it.</span></span> <span data-ttu-id="1df3b-140">Para solicitar que as propriedades de um objeto proxy sejam carregadas, chame o método **load()** no objeto e especifique as propriedades a serem carregadas.</span><span class="sxs-lookup"><span data-stu-id="1df3b-140">To request properties of a proxy object be loaded, call the **load()** method on the object and specify the properties to load.</span></span> 

> [!NOTE]
> <span data-ttu-id="1df3b-141">Se estiver apenas chamando métodos ou definindo propriedades em um objeto proxy, você não precisa chamar o método **load()**.</span><span class="sxs-lookup"><span data-stu-id="1df3b-141">If you are only calling methods or setting properties on a proxy object, you do not need to call the **load()** method.</span></span> <span data-ttu-id="1df3b-142">O método **load()** só é necessário quando você deseja ler propriedades em um objeto proxy.</span><span class="sxs-lookup"><span data-stu-id="1df3b-142">The **load()** method is only required when you want to read properties on a proxy object.</span></span>
 
<span data-ttu-id="1df3b-p112">Assim como as solicitações para definir propriedades ou invocar métodos em objetos proxy, as solicitações para carregar propriedades em objetos proxy são adicionadas à fila de comandos pendentes no contexto de solicitação, sendo executadas na próxima vez que você chamar o método **sync()**. É possível enfileirar quantas chamadas de **load()** forem necessárias no contexto de solicitação.</span><span class="sxs-lookup"><span data-stu-id="1df3b-p112">Just like requests to set properties or invoke methods on proxy objects, requests to load properties on proxy objects get added to the queue of pending commands on the request context, which will run the next time you call the **sync()** method. You can queue up as many **load()** calls on the request context as necessary.</span></span>
 
<span data-ttu-id="1df3b-145">No exemplo a seguir, somente propriedades específicas do intervalo são carregadas.</span><span class="sxs-lookup"><span data-stu-id="1df3b-145">In the following example, only specific properties of the range are loaded.</span></span>
 
```js
Excel.run(function (context) {
  const sheetName = 'Sheet1';
  const rangeAddress = 'A1:B2';
  const myRange = context.workbook.worksheets.getItem(sheetName).getRange(rangeAddress);
 
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
 
<span data-ttu-id="1df3b-146">No exemplo anterior, como `format/font` não é especificado na chamada a **myRange.load()**, a propriedade `format.font.color` não pode ser lida.</span><span class="sxs-lookup"><span data-stu-id="1df3b-146">In the previous example, because `format/font` is not specified in the call to **myRange.load()**, the `format.font.color` property cannot be read.</span></span>

<span data-ttu-id="1df3b-147">Para otimizar o desempenho, você deve especificar explicitamente as propriedades e relações a serem carregadas ao usar o método **load()** em um objeto, conforme abordado em [Otimizações de desempenho da API JavaScript do Excel](performance.md).</span><span class="sxs-lookup"><span data-stu-id="1df3b-147">To optimize performance, you should explicitly specify the properties and relationships to load when using the **load()** method on an object.</span></span> <span data-ttu-id="1df3b-148">Para saber mais sobre o método **load()**, confira os [conceitos avançados da API JavaScript do Excel](excel-add-ins-advanced-concepts.md).</span><span class="sxs-lookup"><span data-stu-id="1df3b-148">For more information about the **load()** method, see [Excel JavaScript API advanced concepts](excel-add-ins-advanced-concepts.md).</span></span>

## <a name="null-or-blank-property-values"></a><span data-ttu-id="1df3b-149">Valores de propriedade nula ou em branco</span><span class="sxs-lookup"><span data-stu-id="1df3b-149">null or blank property values</span></span>
 
### <a name="null-input-in-2-d-array"></a><span data-ttu-id="1df3b-150">entrada nula em uma matriz 2D</span><span class="sxs-lookup"><span data-stu-id="1df3b-150">null input in 2-D Array</span></span>
 
<span data-ttu-id="1df3b-151">No Excel, um intervalo é representado por uma matriz 2D, onde a primeira dimensão é linhas e a segunda dimensão é colunas.</span><span class="sxs-lookup"><span data-stu-id="1df3b-151">In Excel, a range is represented by a 2-D array, where the first dimension is rows and the second dimension is columns.</span></span> <span data-ttu-id="1df3b-152">Para definir valores, o formato do número ou a fórmula apenas para células específicas em um intervalo, especifique os valores, o formato do número ou a fórmula para essas células na matriz 2D, bem como `null` para todas as outras células na matriz 2D.</span><span class="sxs-lookup"><span data-stu-id="1df3b-152">To set values, number format, or formula for only specific cells within a range, specify the values, number format, or formula for those cells in the 2-D array, and specify `null` for all other cells in the 2-D array.</span></span>
 
<span data-ttu-id="1df3b-153">Por exemplo, para atualizar o formato do número apenas para uma célula em um intervalo e manter o formato de número existente para todas as outras células no intervalo, especifique o novo formato de número para a célula a ser atualizada e `null` para todas as outras células.</span><span class="sxs-lookup"><span data-stu-id="1df3b-153">For example, to update the number format for only one cell within a range, and retain the existing number format for all other cells in the range, specify the new number format for the cell to update, and specify `null` for all other cells.</span></span> <span data-ttu-id="1df3b-154">O trecho de código a seguir define um novo formato de número para a quarta célula no intervalo e não altera o formato de número para as primeiras três células no intervalo.</span><span class="sxs-lookup"><span data-stu-id="1df3b-154">The following code snippet sets a new number format for the fourth cell in the range, and leaves the number format unchanged for the first three cells in the range.</span></span>
 
```js
range.values = [['Eurasia', '29.96', '0.25', '15-Feb' ]];
range.numberFormat = [[null, null, null, 'm/d/yyyy;@']];
```
 
### <a name="null-input-for-a-property"></a><span data-ttu-id="1df3b-155">entrada nula para uma propriedade</span><span class="sxs-lookup"><span data-stu-id="1df3b-155">null input for a property</span></span>
 
<span data-ttu-id="1df3b-p116">`null` não é uma entrada válida para uma propriedade única. Por exemplo, o trecho de código a seguir não é válido, pois a propriedade **values** do intervalo não pode ser definida como `null`.</span><span class="sxs-lookup"><span data-stu-id="1df3b-p116">`null` is not a valid input for single property. For example, the following code snippet is not valid, as the **values** property of the range cannot be set to `null`.</span></span>
 
```js
range.values = null;
```
 
<span data-ttu-id="1df3b-158">Da mesma forma, o trecho de código a seguir não é válido, pois `null` não é um valor válido para a propriedade **color**.</span><span class="sxs-lookup"><span data-stu-id="1df3b-158">Likewise, the following code snippet is not valid, as `null` is not a valid value for the **color** property.</span></span>
 
```js
range.format.fill.color =  null;
```
 
### <a name="null-property-values-in-the-response"></a><span data-ttu-id="1df3b-159">Valores da propriedade nula na resposta</span><span class="sxs-lookup"><span data-stu-id="1df3b-159">null property values in the response</span></span>
 
<span data-ttu-id="1df3b-160">A formatação de propriedades como `size` e `color` conterá valores `null` na resposta quando valores diferentes existirem no intervalo especificado.</span><span class="sxs-lookup"><span data-stu-id="1df3b-160">Formatting properties such as `size` and `color` will contain `null` values in the response when different values exist in the specified range.</span></span> <span data-ttu-id="1df3b-161">Por exemplo, se você recuperar um intervalo e carregar sua propriedade `format.font.color`:</span><span class="sxs-lookup"><span data-stu-id="1df3b-161">For example, if you retrieve a range and load its `format.font.color` property:</span></span>
 
* <span data-ttu-id="1df3b-162">Se todas as células no intervalo tiverem a mesma cor de fonte, `range.format.font.color` especificará essa cor.</span><span class="sxs-lookup"><span data-stu-id="1df3b-162">If all cells in the range have the same font color, `range.format.font.color` specifies that color.</span></span>
* <span data-ttu-id="1df3b-163">Se houver várias cores de fonte dentro do intervalo, `range.format.font.color` será `null`.</span><span class="sxs-lookup"><span data-stu-id="1df3b-163">If multiple font colors are present within the range, `range.format.font.color` is `null`.</span></span>
 
### <a name="blank-input-for-a-property"></a><span data-ttu-id="1df3b-164">Entrada em branco para uma propriedade</span><span class="sxs-lookup"><span data-stu-id="1df3b-164">Blank input for a property</span></span>
 
<span data-ttu-id="1df3b-p118">Quando você especificar um valor em branco para uma propriedade (isto é, duas aspas sem espaço entre elas `''`), ele será interpretado como uma instrução para limpar ou redefinir a propriedade. Por exemplo:</span><span class="sxs-lookup"><span data-stu-id="1df3b-p118">When you specify a blank value for a property (i.e., two quotation marks with no space in-between `''`), it will be interpreted as an instruction to clear or reset the property. For example:</span></span>
 
* <span data-ttu-id="1df3b-167">Se você especificar um valor em branco para a propriedade `values` de um intervalo, o conteúdo do intervalo será apagado.</span><span class="sxs-lookup"><span data-stu-id="1df3b-167">If you specify a blank value for the `values` property of a range, the content of the range is cleared.</span></span>
 
* <span data-ttu-id="1df3b-168">Se você especificar um valor em branco para a propriedade `numberFormat`, o formato de número será redefinido para `General`.</span><span class="sxs-lookup"><span data-stu-id="1df3b-168">If you specify a blank value for the `numberFormat` property, the number format is reset to `General`.</span></span>
 
* <span data-ttu-id="1df3b-169">Se você especificar um valor em branco para a propriedade `formula` e a propriedade `formulaLocale`, os valores de fórmula serão apagados.</span><span class="sxs-lookup"><span data-stu-id="1df3b-169">If you specify a blank value for the `formula` property and `formulaLocale` property, the formula values are cleared.</span></span>
 
### <a name="blank-property-values-in-the-response"></a><span data-ttu-id="1df3b-170">Valores da propriedade em branco na resposta</span><span class="sxs-lookup"><span data-stu-id="1df3b-170">Blank property values in the response</span></span>
 
<span data-ttu-id="1df3b-171">Para operações de leitura, um valor de propriedade em branco na resposta (isto é, duas aspas sem espaço entre elas `''`) indica que a célula não contém dados nem valor.</span><span class="sxs-lookup"><span data-stu-id="1df3b-171">For read operations, a blank property value in the response (i.e., two quotation marks with no space in-between `''`) indicates that cell contains no data or value.</span></span> <span data-ttu-id="1df3b-172">No primeiro exemplo abaixo, a primeira e a última célula no intervalo não contêm dados.</span><span class="sxs-lookup"><span data-stu-id="1df3b-172">In the first example below, the first and last cell in the range contain no data.</span></span> <span data-ttu-id="1df3b-173">No segundo exemplo, as primeiras duas células no intervalo não contêm uma fórmula.</span><span class="sxs-lookup"><span data-stu-id="1df3b-173">In the second example, the first two cells in the range do not contain a formula.</span></span>
 
```js
range.values = [['', 'some', 'data', 'in', 'other', 'cells', '']];
```
 
```js
range.formula = [['', '', '=Rand()']];
```
 
## <a name="read-or-write-to-an-unbounded-range"></a><span data-ttu-id="1df3b-174">Ler ou gravar em um intervalo não limitado</span><span class="sxs-lookup"><span data-stu-id="1df3b-174">Read or write to an unbounded range</span></span>
 
### <a name="read-an-unbounded-range"></a><span data-ttu-id="1df3b-175">Ler um intervalo não limitado</span><span class="sxs-lookup"><span data-stu-id="1df3b-175">Read an unbounded range</span></span>
 
<span data-ttu-id="1df3b-p120">Um endereço de intervalo não limitado é um endereço de intervalo que especifica colunas ou linhas inteiras. Por exemplo:</span><span class="sxs-lookup"><span data-stu-id="1df3b-p120">An unbounded range address is a range address that specifies either entire column(s) or entire row(s). For example:</span></span>
 
* <span data-ttu-id="1df3b-178">Endereços de intervalo composto por colunas inteiras:</span><span class="sxs-lookup"><span data-stu-id="1df3b-178">Range addresses comprised of entire column(s):</span></span><ul><li>`C:C`</li><li>`A:F`</li></ul>
* <span data-ttu-id="1df3b-179">Endereços de intervalo composto por linhas inteiras:</span><span class="sxs-lookup"><span data-stu-id="1df3b-179">Range addresses comprised of entire row(s):</span></span><ul><li>`2:2`</li><li>`1:4`</li></ul>
 
<span data-ttu-id="1df3b-180">Quando uma API faz uma solicitação para recuperar um intervalo não limitado (por exemplo, `getRange('C:C')`), a resposta conterá valores `null` para as propriedades no nível de célula, como `values`, `text`, `numberFormat` e `formula`.</span><span class="sxs-lookup"><span data-stu-id="1df3b-180">When the API makes a request to retrieve an unbounded range (for example, `getRange('C:C')`), the response will contain `null` values for cell-level properties such as `values`, `text`, `numberFormat`, and `formula`.</span></span> <span data-ttu-id="1df3b-181">Outras propriedades do intervalo, como `address` e `cellCount`, conterão valores válidos para o intervalo não limitado.</span><span class="sxs-lookup"><span data-stu-id="1df3b-181">Other properties of the range, such as `address` and `cellCount`, will contain valid values for the unbounded range.</span></span>
 
### <a name="write-to-an-unbounded-range"></a><span data-ttu-id="1df3b-182">Gravar em um intervalo não limitado</span><span class="sxs-lookup"><span data-stu-id="1df3b-182">Write to an unbounded range</span></span>
 
<span data-ttu-id="1df3b-183">Não é possível definir propriedades no nível de célula, como `values`, `numberFormat` e `formula`, no intervalo não limitado, pois a solicitação de entrada é muito grande.</span><span class="sxs-lookup"><span data-stu-id="1df3b-183">You cannot set cell-level properties such as `values`, `numberFormat`, and `formula` on unbounded range because the input request is too large.</span></span> <span data-ttu-id="1df3b-184">Por exemplo, o trecho de código a seguir não é válida porque ele tenta especificar `values` para um intervalo não limitado.</span><span class="sxs-lookup"><span data-stu-id="1df3b-184">For example, the following code snippet is not valid because it attempts to specify `values` for an unbounded range.</span></span> <span data-ttu-id="1df3b-185">A API retornará um erro se você tentar definir as propriedades no nível de célula para um intervalo não limitado.</span><span class="sxs-lookup"><span data-stu-id="1df3b-185">The API will return an error if you attempt to set cell-level properties for an unbounded range.</span></span>
 
```js
const range = context.workbook.worksheets.getActiveWorksheet().getRange('A:B');
range.values = 'Due Date';
```
 
## <a name="read-or-write-to-a-large-range"></a><span data-ttu-id="1df3b-186">Ler ou gravar em um intervalo grande</span><span class="sxs-lookup"><span data-stu-id="1df3b-186">Read or write to a large range</span></span>
 
<span data-ttu-id="1df3b-187">Se um intervalo contiver um grande número de células, valores, formatos de número e/ou fórmulas, talvez não seja possível executar operações de API nesse intervalo.</span><span class="sxs-lookup"><span data-stu-id="1df3b-187">If a range contains a large number of cells, values, number formats, and/or formulas, it may not be possible to run API operations on that range.</span></span> <span data-ttu-id="1df3b-188">A API sempre fará a melhor tentativa de executar a operação solicitada em um intervalo (isto é, para recuperar ou gravar os dados especificados), mas tentar executar operações de leitura ou gravação para um intervalo grande pode resultar em um erro de API devido à utilização excessiva de recursos.</span><span class="sxs-lookup"><span data-stu-id="1df3b-188">The API will always make a best attempt to run the requested operation on a range (i.e., to retrieve or write the specified data), but attempting to perform read or write operations for a large range may result in an API error due to excessive resource utilization.</span></span> <span data-ttu-id="1df3b-189">Para evitar tais erros, é recomendável executar operações de leitura ou gravação separadas para subconjuntos menores de um intervalo grande, em vez de tentar executar uma única operação de leitura ou gravação em um intervalo grande.</span><span class="sxs-lookup"><span data-stu-id="1df3b-189">To avoid such errors, we recommend that you run separate read or write operations for smaller subsets of a large range, instead of attempting to run a single read or write operation on a large range.</span></span>
 
## <a name="update-all-cells-in-a-range"></a><span data-ttu-id="1df3b-190">Atualizar todas as células em um intervalo</span><span class="sxs-lookup"><span data-stu-id="1df3b-190">Update all cells in a range</span></span>
 
<span data-ttu-id="1df3b-191">Para aplicar a mesma atualização a todas as células em um intervalo, (por exemplo, para popular todas as células com o mesmo valor, definir o mesmo formato de número ou popular todas as células com a mesma fórmula), defina a propriedade correspondente no objeto **range** para o valor (único) desejado.</span><span class="sxs-lookup"><span data-stu-id="1df3b-191">To apply the same update to all cells in a range, (for example, to populate all cells with the same value, set the same number format, or populate all cells with the same formula), set the corresponding property on the **range** object to the desired (single) value.</span></span>
 
<span data-ttu-id="1df3b-192">O exemplo a seguir obtém um intervalo que contém 20 células e, em seguida, define o formato de número e popula todas as células do intervalo com o valor **11/3/2015**.</span><span class="sxs-lookup"><span data-stu-id="1df3b-192">The following example gets a range that contains 20 cells, and then sets the number format and populates all cells in the range with the value **3/11/2015**.</span></span>
 
```js
Excel.run(function (context) {
  const sheetName = 'Sheet1';
  const rangeAddress = 'A1:A20';
  const worksheet = context.workbook.worksheets.getItem(sheetName);
 
  const range = worksheet.getRange(rangeAddress);
  range.numberFormat = 'm/d/yyyy';
  range.values = '3/11/2015';
  range.load('text');
 
  return context.sync()
    .then(function () {
      console.log(range.text);
  });
}).catch(function (error) {
  console.log('Error: ' + error);
  if (error instanceof OfficeExtension.Error) {
    console.log('Debug info: ' + JSON.stringify(error.debugInfo));
  }
});
```
 
## <a name="error-messages"></a><span data-ttu-id="1df3b-193">Mensagens de erro</span><span class="sxs-lookup"><span data-stu-id="1df3b-193">Error messages</span></span>
 
<span data-ttu-id="1df3b-194">Quando ocorrer um erro de API, a API retornará um objeto **error** que contém um código e uma mensagem.</span><span class="sxs-lookup"><span data-stu-id="1df3b-194">When an API error occurs, the API will return an **error** object that contains a code and a message.</span></span> <span data-ttu-id="1df3b-195">A tabela a seguir define uma lista de erros que a API pode retornar.</span><span class="sxs-lookup"><span data-stu-id="1df3b-195">The following table defines a list of errors that the API may return.</span></span>
 
|<span data-ttu-id="1df3b-196">error.code</span><span class="sxs-lookup"><span data-stu-id="1df3b-196">error.code</span></span> | <span data-ttu-id="1df3b-197">error.message</span><span class="sxs-lookup"><span data-stu-id="1df3b-197">error.message</span></span> |
|:----------|:--------------|
|<span data-ttu-id="1df3b-198">InvalidArgument</span><span class="sxs-lookup"><span data-stu-id="1df3b-198">InvalidArgument</span></span> |<span data-ttu-id="1df3b-199">O argumento é inválido, está ausente ou tem um formato incorreto.</span><span class="sxs-lookup"><span data-stu-id="1df3b-199">The argument is invalid or missing or has an incorrect format.</span></span>|
|<span data-ttu-id="1df3b-200">InvalidRequest</span><span class="sxs-lookup"><span data-stu-id="1df3b-200">InvalidRequest</span></span>  |<span data-ttu-id="1df3b-201">Não é possível processar a solicitação.</span><span class="sxs-lookup"><span data-stu-id="1df3b-201">Cannot process the request.</span></span>|
|<span data-ttu-id="1df3b-202">InvalidReference</span><span class="sxs-lookup"><span data-stu-id="1df3b-202">InvalidReference</span></span>|<span data-ttu-id="1df3b-203">Esta referência não é válida para a operação atual.</span><span class="sxs-lookup"><span data-stu-id="1df3b-203">This reference is not valid for the current operation.</span></span>|
|<span data-ttu-id="1df3b-204">InvalidBinding</span><span class="sxs-lookup"><span data-stu-id="1df3b-204">InvalidBinding</span></span>  |<span data-ttu-id="1df3b-205">Esta associação de objetos não é mais válida devido às atualizações anteriores.</span><span class="sxs-lookup"><span data-stu-id="1df3b-205">This object binding is no longer valid due to previous updates.</span></span>|
|<span data-ttu-id="1df3b-206">InvalidSelection</span><span class="sxs-lookup"><span data-stu-id="1df3b-206">InvalidSelection</span></span>|<span data-ttu-id="1df3b-207">A seleção atual é inválida para esta operação.</span><span class="sxs-lookup"><span data-stu-id="1df3b-207">The current selection is invalid for this operation.</span></span>|
|<span data-ttu-id="1df3b-208">Unauthenticated</span><span class="sxs-lookup"><span data-stu-id="1df3b-208">Unauthenticated</span></span> |<span data-ttu-id="1df3b-209">Informações de autenticação necessárias estão ausentes ou inválidas.</span><span class="sxs-lookup"><span data-stu-id="1df3b-209">Required authentication information is either missing or invalid.</span></span>|
|<span data-ttu-id="1df3b-210">AccessDenied</span><span class="sxs-lookup"><span data-stu-id="1df3b-210">AccessDenied</span></span> |<span data-ttu-id="1df3b-211">Você não pode realizar a operação solicitada.</span><span class="sxs-lookup"><span data-stu-id="1df3b-211">You cannot perform the requested operation.</span></span>|
|<span data-ttu-id="1df3b-212">ItemNotFound</span><span class="sxs-lookup"><span data-stu-id="1df3b-212">ItemNotFound</span></span> |<span data-ttu-id="1df3b-213">O recurso solicitado não existe.</span><span class="sxs-lookup"><span data-stu-id="1df3b-213">The requested resource doesn't exist.</span></span>|
|<span data-ttu-id="1df3b-214">ActivityLimitReached</span><span class="sxs-lookup"><span data-stu-id="1df3b-214">ActivityLimitReached</span></span>|<span data-ttu-id="1df3b-215">O limite de atividades foi alcançado.</span><span class="sxs-lookup"><span data-stu-id="1df3b-215">Activity limit has been reached.</span></span>|
|<span data-ttu-id="1df3b-216">GeneralException</span><span class="sxs-lookup"><span data-stu-id="1df3b-216">GeneralException</span></span>|<span data-ttu-id="1df3b-217">Ocorreu um erro interno ao processar a solicitação.</span><span class="sxs-lookup"><span data-stu-id="1df3b-217">There was an internal error while processing the request.</span></span>|
|<span data-ttu-id="1df3b-218">NotImplemented</span><span class="sxs-lookup"><span data-stu-id="1df3b-218">NotImplemented</span></span>  |<span data-ttu-id="1df3b-219">O recurso solicitado não foi implementado.</span><span class="sxs-lookup"><span data-stu-id="1df3b-219">The requested feature isn't implemented.</span></span>|
|<span data-ttu-id="1df3b-220">ServiceNotAvailable</span><span class="sxs-lookup"><span data-stu-id="1df3b-220">ServiceNotAvailable</span></span>|<span data-ttu-id="1df3b-221">O serviço não está disponível.</span><span class="sxs-lookup"><span data-stu-id="1df3b-221">The service is unavailable.</span></span>|
|<span data-ttu-id="1df3b-222">Conflito</span><span class="sxs-lookup"><span data-stu-id="1df3b-222">Conflict</span></span>              |<span data-ttu-id="1df3b-223">A solicitação não pôde ser processada devido a um conflito.</span><span class="sxs-lookup"><span data-stu-id="1df3b-223">Request could not be processed because of a conflict.</span></span>|
|<span data-ttu-id="1df3b-224">ItemAlreadyExists</span><span class="sxs-lookup"><span data-stu-id="1df3b-224">ItemAlreadyExists</span></span>|<span data-ttu-id="1df3b-225">O recurso que está sendo criado já existe.</span><span class="sxs-lookup"><span data-stu-id="1df3b-225">The resource being created already exists.</span></span>|
|<span data-ttu-id="1df3b-226">UnsupportedOperation</span><span class="sxs-lookup"><span data-stu-id="1df3b-226">UnsupportedOperation</span></span>|<span data-ttu-id="1df3b-227">Não há suporte para a operação que está sendo tentada.</span><span class="sxs-lookup"><span data-stu-id="1df3b-227">The operation being attempted is not supported.</span></span>|
|<span data-ttu-id="1df3b-228">RequestAborted</span><span class="sxs-lookup"><span data-stu-id="1df3b-228">RequestAborted</span></span>|<span data-ttu-id="1df3b-229">A solicitação foi anulada durante o tempo de execução.</span><span class="sxs-lookup"><span data-stu-id="1df3b-229">The request was aborted during run time.</span></span>|
|<span data-ttu-id="1df3b-230">ApiNotAvailable</span><span class="sxs-lookup"><span data-stu-id="1df3b-230">ApiNotAvailable</span></span>|<span data-ttu-id="1df3b-231">A API solicitada não está disponível.</span><span class="sxs-lookup"><span data-stu-id="1df3b-231">The requested API is not available.</span></span>|
|<span data-ttu-id="1df3b-232">InsertDeleteConflict</span><span class="sxs-lookup"><span data-stu-id="1df3b-232">InsertDeleteConflict</span></span>|<span data-ttu-id="1df3b-233">A tentativa de operação de exclusão ou inserção resultou em um conflito.</span><span class="sxs-lookup"><span data-stu-id="1df3b-233">The insert or delete operation attempted resulted in a conflict.</span></span>|
|<span data-ttu-id="1df3b-234">InvalidOperation</span><span class="sxs-lookup"><span data-stu-id="1df3b-234">InvalidOperation</span></span>|<span data-ttu-id="1df3b-235">A tentativa de operação é inválida no objeto.</span><span class="sxs-lookup"><span data-stu-id="1df3b-235">The operation attempted is invalid on the object.</span></span>|
 
## <a name="see-also"></a><span data-ttu-id="1df3b-236">Veja também</span><span class="sxs-lookup"><span data-stu-id="1df3b-236">See also</span></span>
 
* [<span data-ttu-id="1df3b-237">Introdução aos suplementos do Excel</span><span class="sxs-lookup"><span data-stu-id="1df3b-237">Get started with Excel add-ins</span></span>](excel-add-ins-get-started-overview.md)
* [<span data-ttu-id="1df3b-238">Exemplos de código de suplementos do Excel</span><span class="sxs-lookup"><span data-stu-id="1df3b-238">Excel add-ins code samples</span></span>](https://developer.microsoft.com/office/gallery/?filterBy=Samples)
* [<span data-ttu-id="1df3b-239">Otimização de desempenho da API JavaScript do Excel</span><span class="sxs-lookup"><span data-stu-id="1df3b-239">Excel JavaScript API performance optimization</span></span>](https://docs.microsoft.com/office/dev/add-ins/excel/performance)
* [<span data-ttu-id="1df3b-240">Referência da API JavaScript do Excel</span><span class="sxs-lookup"><span data-stu-id="1df3b-240">Excel JavaScript API reference</span></span>](https://dev.office.com/reference/add-ins/excel/excel-add-ins-reference-overview)
