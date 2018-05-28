---
title: Principais conceitos da API JavaScript do Excel
description: ''
ms.date: 12/04/2017
ms.openlocfilehash: 1582268a3bdac2b7fe63c4b0a48cf1a19f85bd31
ms.sourcegitcommit: c72c35e8389c47a795afbac1b2bcf98c8e216d82
ms.translationtype: HT
ms.contentlocale: pt-BR
ms.lasthandoff: 05/23/2018
---
# <a name="excel-javascript-api-core-concepts"></a><span data-ttu-id="ea7f4-102">Principais conceitos da API JavaScript do Excel</span><span class="sxs-lookup"><span data-stu-id="ea7f4-102">Excel JavaScript API core concepts</span></span>
 
<span data-ttu-id="ea7f4-103">Este artigo descreve como usar a [API JavaScript do Excel](https://dev.office.com/reference/add-ins/excel/excel-add-ins-reference-overview) para desenvolver suplementos para o Excel 2016.</span><span class="sxs-lookup"><span data-stu-id="ea7f4-103">This article describes how to use the [Excel JavaScript API](https://dev.office.com/reference/add-ins/excel/excel-add-ins-reference-overview) to build add-ins for Excel 2016.</span></span> <span data-ttu-id="ea7f4-104">Ele apresenta os conceitos b?sicos que s?o fundamentais para usar a API e fornece orienta??es para executar tarefas espec?ficas, como leitura ou grava??o em um intervalo grande, atualiza??o de todas as c?lulas do intervalo e muito mais.</span><span class="sxs-lookup"><span data-stu-id="ea7f4-104">It introduces core concepts that are fundamental to using the API and provides guidance for performing specific tasks such as reading or writing to a large range, updating all cells in range, and more.</span></span>

## <a name="asynchronous-nature-of-excel-apis"></a><span data-ttu-id="ea7f4-105">Natureza ass?ncrona das APIs do Excel</span><span class="sxs-lookup"><span data-stu-id="ea7f4-105">Asynchronous nature of Excel APIs</span></span>

<span data-ttu-id="ea7f4-106">Os suplementos do Excel baseados na Web s?o executados dentro de um cont?iner de navegador que ? inserido no aplicativo do Office em plataformas baseadas em desktop, como Office para Windows, e executado dentro de um iFrame HTML no Office Online.</span><span class="sxs-lookup"><span data-stu-id="ea7f4-106">The web-based Excel add-ins run inside a browser container that is embedded within the Office application on desktop-based platforms such as Office for Windows and runs inside an HTML iFrame in Office Online.</span></span> <span data-ttu-id="ea7f4-107">N?o ? poss?vel habilitar a API Office.js para interagir de modo s?ncrono com o host do Excel em todas as plataformas suportadas devido ?s considera??es de desempenho.</span><span class="sxs-lookup"><span data-stu-id="ea7f4-107">Enabling the Office.js API to interact synchronously with the Excel host across all supported platforms is not feasible due to performance considerations.</span></span> <span data-ttu-id="ea7f4-108">Desse modo, a chamada ? API **sync()** na Office.js retorna uma [promessa](https://developer.mozilla.org/en-US/docs/Web/JavaScript/Reference/Global_Objects/Promise) que ? resolvida quando o aplicativo Excel conclui as a??es solicitadas de leitura ou grava??o.</span><span class="sxs-lookup"><span data-stu-id="ea7f4-108">Therefore, the **sync()** API call in Office.js returns a [promise](https://developer.mozilla.org/en-US/docs/Web/JavaScript/Reference/Global_Objects/Promise) that is resolved when the Excel application completes the requested read or write actions.</span></span> <span data-ttu-id="ea7f4-109">Al?m disso, voc? pode enfileirar v?rias a??es, como configurar propriedades ou invocar m?todos, e execut?-las como um lote de comandos com uma ?nica chamada a **sync()**, em vez de enviar uma solicita??o separada para cada a??o.</span><span class="sxs-lookup"><span data-stu-id="ea7f4-109">Also, you can queue up multiple actions, such as setting properties or invoking methods, and run them as a batch of commands with a single call to **sync()**, rather than sending a separate request for each action.</span></span> <span data-ttu-id="ea7f4-110">As se??es a seguir descrevem como fazer isso usando as APIs **Excel.run()** e **sync()**.</span><span class="sxs-lookup"><span data-stu-id="ea7f4-110">The following sections describe how to accomplish this using the **Excel.run()** and **sync()** APIs.</span></span>
 
## <a name="excelrun"></a><span data-ttu-id="ea7f4-111">Excel.run</span><span class="sxs-lookup"><span data-stu-id="ea7f4-111">Excel.run</span></span>
 
<span data-ttu-id="ea7f4-112">A **Excel.run** executa uma fun??o em que voc? especifica as a??es a serem executadas no modelo de objeto do Excel.</span><span class="sxs-lookup"><span data-stu-id="ea7f4-112">**Excel.run** executes a function where you specify the actions to perform against the Excel object model.</span></span> <span data-ttu-id="ea7f4-113">A **Excel.run** cria automaticamente um contexto de solicita??o que pode ser usado para sua intera??o com os objetos do Excel.</span><span class="sxs-lookup"><span data-stu-id="ea7f4-113">**Excel.run** automatically creates a request context that you can use to interact with Excel objects.</span></span> <span data-ttu-id="ea7f4-114">Quando a **Excel.run** ? conclu?da, uma promessa ? resolvida e todos os objetos que foram alocados em tempo de execu??o s?o lan?ados automaticamente.</span><span class="sxs-lookup"><span data-stu-id="ea7f4-114">When **Excel.run** completes, a promise is resolved, and any objects that were allocated at runtime are automatically released.</span></span>
 
<span data-ttu-id="ea7f4-115">O exemplo a seguir mostra como usar a **Excel.run**.</span><span class="sxs-lookup"><span data-stu-id="ea7f4-115">The following example shows how to use **Excel.run**.</span></span> <span data-ttu-id="ea7f4-116">A instru??o catch captura e grava em log os erros que ocorrem na **Excel.run**.</span><span class="sxs-lookup"><span data-stu-id="ea7f4-116">The catch statement catches and logs errors that occur within the **Excel.run**.</span></span>
 
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

## <a name="request-context"></a><span data-ttu-id="ea7f4-117">Contexto de solicita??o</span><span class="sxs-lookup"><span data-stu-id="ea7f4-117">Request context</span></span>
 
<span data-ttu-id="ea7f4-p105">O Excel e seu suplemento s?o executados em dois processos diferentes. Como eles usam diferentes ambientes de tempo de execu??o, os suplementos do Excel exigem um objeto **RequestContext** para conectar o suplemento aos objetos no Excel, como planilhas, intervalos, gr?ficos e tabelas.</span><span class="sxs-lookup"><span data-stu-id="ea7f4-p105">Excel and your add-in run in two different processes. Since they use different runtime environments, Excel add-ins require a **RequestContext** object in order to connect your add-in to objects in Excel such as worksheets, ranges, charts, and tables.</span></span>
 
## <a name="proxy-objects"></a><span data-ttu-id="ea7f4-120">Objetos proxy</span><span class="sxs-lookup"><span data-stu-id="ea7f4-120">Proxy objects</span></span>
 
<span data-ttu-id="ea7f4-121">Os objetos JavaScript do Excel que voc? declara e usa em um suplemento s?o objetos proxy.</span><span class="sxs-lookup"><span data-stu-id="ea7f4-121">The Excel JavaScript objects that you declare and use in an add-in are proxy objects.</span></span> <span data-ttu-id="ea7f4-122">Todos os m?todos invocados, ou as propriedades definidas ou carregadas em objetos proxy s?o simplesmente adicionados a uma fila de comandos pendentes.</span><span class="sxs-lookup"><span data-stu-id="ea7f4-122">Any methods that you invoke or properties that you set or load on proxy objects are simply added to a queue of pending commands.</span></span> <span data-ttu-id="ea7f4-123">Quando voc? chama o m?todo **sync()** no contexto de solicita??o (por exemplo, `context.sync()`), os comandos enfileirados s?o expedidos para o Excel e executados.</span><span class="sxs-lookup"><span data-stu-id="ea7f4-123">When you call the **sync()** method on the request context (for example, `context.sync()`), the queued commands are dispatched to Excel and run.</span></span> <span data-ttu-id="ea7f4-124">A API JavaScript do Excel ? basicamente centrada em lote.</span><span class="sxs-lookup"><span data-stu-id="ea7f4-124">The Excel JavaScript API is fundamentally batch-centric.</span></span> <span data-ttu-id="ea7f4-125">Voc? pode enfileirar quantas altera??es desejar no contexto de solicita??o e depois chamar o m?todo **sync()** para executar o lote de comandos enfileirados.</span><span class="sxs-lookup"><span data-stu-id="ea7f4-125">You can queue up as many changes as you wish on the request context, and then call the **sync()** method to run the batch of queued commands.</span></span>
 
<span data-ttu-id="ea7f4-126">Por exemplo, o trecho de c?digo a seguir declara o objeto JavaScript local **selectedRange** para fazer refer?ncia a um intervalo selecionado no documento do Excel e, em seguida, define algumas propriedades nesse objeto.</span><span class="sxs-lookup"><span data-stu-id="ea7f4-126">For example, the following code snippet declares the local JavaScript object **selectedRange** to reference a selected range in the Excel document, and then sets some properties on that object.</span></span> <span data-ttu-id="ea7f4-127">O objeto **selectedRange** ? um objeto proxy, de modo que as propriedades que s?o definidas e o m?todo que ? invocado nesse objeto n?o ser?o refletidos no documento do Excel at? que o suplemento chame **context.sync()**.</span><span class="sxs-lookup"><span data-stu-id="ea7f4-127">The **selectedRange** object is a proxy object, so the properties that are set and method that is invoked on that object will not be reflected in the Excel document until your add-in calls **context.sync()**.</span></span>
 
```js
const selectedRange = context.workbook.getSelectedRange();
selectedRange.format.fill.color = "#4472C4";
selectedRange.format.font.color = "white";
selectedRange.format.autofitColumns();
```
 
### <a name="sync"></a><span data-ttu-id="ea7f4-128">sync()</span><span class="sxs-lookup"><span data-stu-id="ea7f4-128">sync()</span></span>
 
<span data-ttu-id="ea7f4-129">Chamar o m?todo **sync()** no contexto de solicita??o sincroniza o estado entre objetos proxy e objetos no documento do Excel.</span><span class="sxs-lookup"><span data-stu-id="ea7f4-129">Calling the **sync()** method on the request context synchronizes the state between proxy objects and objects in the Excel document.</span></span> <span data-ttu-id="ea7f4-130">O m?todo **sync()** executa todos os comandos que s?o enfileirados no contexto de solicita??o e recupera valores para qualquer propriedade que deva ser carregada nos objetos proxy.</span><span class="sxs-lookup"><span data-stu-id="ea7f4-130">The **sync()** method runs any commands that are queued on the request context and retrieves values for any properties that should be loaded on the proxy objects.</span></span> <span data-ttu-id="ea7f4-131">O m?todo **sync()** ? executado de modo ass?ncrono e retorna uma [promessa](https://developer.mozilla.org/en-US/docs/Web/JavaScript/Reference/Global_Objects/Promise), que ? resolvida quando o m?todo **sync()** ? conclu?do.</span><span class="sxs-lookup"><span data-stu-id="ea7f4-131">The **sync()** method executes asynchronously and returns a [promise](https://developer.mozilla.org/en-US/docs/Web/JavaScript/Reference/Global_Objects/Promise), which is resolved when the **sync()** method completes.</span></span>
 
<span data-ttu-id="ea7f4-132">O exemplo a seguir mostra uma fun??o de lote que define um objeto proxy JavaScript local (**selectedRange**), carrega uma propriedade desse objeto e, em seguida, usa o padr?o Promessas do JavaScript para chamar **context.sync()** a fim de sincronizar o estado entre objetos proxy e objetos no documento do Excel.</span><span class="sxs-lookup"><span data-stu-id="ea7f4-132">The following example shows a batch function that defines a local JavaScript proxy object (**selectedRange**), loads a property of that object, and then uses the JavaScript Promises pattern to call **context.sync()** to synchronize the state between proxy objects and objects in the Excel document.</span></span>
 
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
 
<span data-ttu-id="ea7f4-133">No exemplo anterior, **selectedRange** est? definido e sua propriedade **address** ? carregada quando **context.sync()** ? chamado.</span><span class="sxs-lookup"><span data-stu-id="ea7f4-133">In the previous example, **selectedRange** is set and its **address** property is loaded when **context.sync()** is called.</span></span>
 
<span data-ttu-id="ea7f4-134">Como **sync()** ? uma opera??o ass?ncrona que retorna uma promessa, voc? sempre deve **retornar** a promessa (no JavaScript).</span><span class="sxs-lookup"><span data-stu-id="ea7f4-134">Because **sync()** is an asynchronous operation that returns a promise, you should always **return** the promise (in JavaScript).</span></span> <span data-ttu-id="ea7f4-135">Isso garante que a opera??o **sync()** seja conclu?da antes que o script continue sendo executado.</span><span class="sxs-lookup"><span data-stu-id="ea7f4-135">Doing so ensures that the **sync()** operation completes before the script continues to run.</span></span> <span data-ttu-id="ea7f4-136">Para obter mais informa??es sobre como otimizar o desempenho com **sync()**, confira [Otimiza??o de desempenho da API JavaScript do Excel](https://dev.office.com/reference/add-ins/excel/performance.md).</span><span class="sxs-lookup"><span data-stu-id="ea7f4-136">For more information about optimizing performance with **sync()**, see [Excel JavaScript API performance optimization](https://dev.office.com/reference/add-ins/excel/performance.md).</span></span>
 
### <a name="load"></a><span data-ttu-id="ea7f4-137">load()</span><span class="sxs-lookup"><span data-stu-id="ea7f4-137">load()</span></span>
 
<span data-ttu-id="ea7f4-138">Para que voc? possa ler as propriedades de um objeto proxy, ? preciso carregar explicitamente as propriedades para popular o objeto proxy com dados do documento do Excel e chamar **context.sync()**.</span><span class="sxs-lookup"><span data-stu-id="ea7f4-138">Before you can read the properties of a proxy object, you must explicitly load the properties to populate the proxy object with data from the Excel document, and then call **context.sync()**.</span></span> <span data-ttu-id="ea7f4-139">Por exemplo, se voc? criar um objeto proxy para fazer refer?ncia a um intervalo selecionado e, em seguida, quiser ler a propriedade **address** do intervalo selecionado, ser? preciso carregar a propriedade **address** para que seja poss?vel l?-la.</span><span class="sxs-lookup"><span data-stu-id="ea7f4-139">For example, if you create a proxy object to reference a selected range, and then want to read the selected range's **address** property, you need to load the **address** property before you can read it.</span></span> <span data-ttu-id="ea7f4-140">Para solicitar que as propriedades de um objeto proxy sejam carregadas, chame o m?todo **load()** no objeto e especifique as propriedades a serem carregadas.</span><span class="sxs-lookup"><span data-stu-id="ea7f4-140">To request properties of a proxy object be loaded, call the **load()** method on the object and specify the properties to load.</span></span> 

> [!NOTE]
> <span data-ttu-id="ea7f4-141">Se estiver apenas chamando m?todos ou definindo propriedades em um objeto proxy, voc? n?o precisa chamar o m?todo **load()**.</span><span class="sxs-lookup"><span data-stu-id="ea7f4-141">If you are only calling methods or setting properties on a proxy object, you do not need to call the **load()** method.</span></span> <span data-ttu-id="ea7f4-142">O m?todo **load()** s? ? necess?rio quando voc? deseja ler propriedades em um objeto proxy.</span><span class="sxs-lookup"><span data-stu-id="ea7f4-142">The **load()** method is only required when you want to read properties on a proxy object.</span></span>
 
<span data-ttu-id="ea7f4-p112">Assim como as solicita??es para definir propriedades ou invocar m?todos em objetos proxy, as solicita??es para carregar propriedades em objetos proxy s?o adicionadas ? fila de comandos pendentes no contexto de solicita??o, sendo executadas na pr?xima vez que voc? chamar o m?todo **sync()**. ? poss?vel enfileirar quantas chamadas de **load()** forem necess?rias no contexto de solicita??o.</span><span class="sxs-lookup"><span data-stu-id="ea7f4-p112">Just like requests to set properties or invoke methods on proxy objects, requests to load properties on proxy objects get added to the queue of pending commands on the request context, which will run the next time you call the **sync()** method. You can queue up as many **load()** calls on the request context as necessary.</span></span>
 
<span data-ttu-id="ea7f4-145">No exemplo a seguir, somente propriedades espec?ficas do intervalo s?o carregadas.</span><span class="sxs-lookup"><span data-stu-id="ea7f4-145">In the following example, only specific properties of the range are loaded.</span></span>
 
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
 
<span data-ttu-id="ea7f4-146">No exemplo anterior, como `format/font` n?o ? especificado na chamada a **myRange.load()**, a propriedade `format.font.color` n?o pode ser lida.</span><span class="sxs-lookup"><span data-stu-id="ea7f4-146">In the previous example, because `format/font` is not specified in the call to **myRange.load()**, the `format.font.color` property cannot be read.</span></span>

<span data-ttu-id="ea7f4-147">Para otimizar o desempenho, voc? deve especificar explicitamente as propriedades e os relacionamentos a serem carregados ao usar o m?todo **load()** em um objeto, conforme [Otimiza??es de desempenho da API JavaScript do Excel](performance.md).</span><span class="sxs-lookup"><span data-stu-id="ea7f4-147">To optimize performance, you should explicitly specify the properties and relationships to load when using the **load()** method on an object.</span></span> <span data-ttu-id="ea7f4-148">Para saber mais sobre o m?todo **load()**, confira os [conceitos avan?ados da API JavaScript do Excel](excel-add-ins-advanced-concepts.md).</span><span class="sxs-lookup"><span data-stu-id="ea7f4-148">For more information about the **load()** method, see [Excel JavaScript API advanced concepts](excel-add-ins-advanced-concepts.md).</span></span>

## <a name="null-or-blank-property-values"></a><span data-ttu-id="ea7f4-149">Valores de propriedade nula ou em branco</span><span class="sxs-lookup"><span data-stu-id="ea7f4-149">null or blank property values</span></span>
 
### <a name="null-input-in-2-d-array"></a><span data-ttu-id="ea7f4-150">entrada nula em uma matriz 2D</span><span class="sxs-lookup"><span data-stu-id="ea7f4-150">null input in 2-D Array</span></span>
 
<span data-ttu-id="ea7f4-151">No Excel, um intervalo ? representado por uma matriz 2D, onde a primeira dimens?o ? linhas e a segunda dimens?o ? colunas.</span><span class="sxs-lookup"><span data-stu-id="ea7f4-151">In Excel, a range is represented by a 2-D array, where the first dimension is rows and the second dimension is columns.</span></span> <span data-ttu-id="ea7f4-152">Para definir valores, o formato do n?mero ou a f?rmula apenas para c?lulas espec?ficas em um intervalo, especifique os valores, o formato do n?mero ou a f?rmula para essas c?lulas na matriz 2D, bem como `null` para todas as outras c?lulas na matriz 2D.</span><span class="sxs-lookup"><span data-stu-id="ea7f4-152">To set values, number format, or formula for only specific cells within a range, specify the values, number format, or formula for those cells in the 2-D array, and specify `null` for all other cells in the 2-D array.</span></span>
 
<span data-ttu-id="ea7f4-153">Por exemplo, para atualizar o formato do n?mero apenas para uma c?lula em um intervalo e manter o formato de n?mero existente para todas as outras c?lulas no intervalo, especifique o novo formato de n?mero para a c?lula a ser atualizada e `null` para todas as outras c?lulas.</span><span class="sxs-lookup"><span data-stu-id="ea7f4-153">For example, to update the number format for only one cell within a range, and retain the existing number format for all other cells in the range, specify the new number format for the cell to update, and specify `null` for all other cells.</span></span> <span data-ttu-id="ea7f4-154">O trecho de c?digo a seguir define um novo formato de n?mero para a quarta c?lula no intervalo e n?o altera o formato de n?mero para as primeiras tr?s c?lulas no intervalo.</span><span class="sxs-lookup"><span data-stu-id="ea7f4-154">The following code snippet sets a new number format for the fourth cell in the range, and leaves the number format unchanged for the first three cells in the range.</span></span>
 
```js
range.values = [['Eurasia', '29.96', '0.25', '15-Feb' ]];
range.numberFormat = [[null, null, null, 'm/d/yyyy;@']];
```
 
### <a name="null-input-for-a-property"></a><span data-ttu-id="ea7f4-155">entrada nula para uma propriedade</span><span class="sxs-lookup"><span data-stu-id="ea7f4-155">null input for a property</span></span>
 
<span data-ttu-id="ea7f4-p116">`null` n?o ? uma entrada v?lida para uma propriedade ?nica. Por exemplo, o trecho de c?digo a seguir n?o ? v?lido, pois a propriedade **values** do intervalo n?o pode ser definida como `null`.</span><span class="sxs-lookup"><span data-stu-id="ea7f4-p116">`null` is not a valid input for single property. For example, the following code snippet is not valid, as the **values** property of the range cannot be set to `null`.</span></span>
 
```js
range.values = null;
```
 
<span data-ttu-id="ea7f4-158">Da mesma forma, o trecho de c?digo a seguir n?o ? v?lido, pois `null` n?o ? um valor v?lido para a propriedade **color**.</span><span class="sxs-lookup"><span data-stu-id="ea7f4-158">Likewise, the following code snippet is not valid, as `null` is not a valid value for the **color** property.</span></span>
 
```js
range.format.fill.color =  null;
```
 
### <a name="null-property-values-in-the-response"></a><span data-ttu-id="ea7f4-159">Valores da propriedade nula na resposta</span><span class="sxs-lookup"><span data-stu-id="ea7f4-159">null property values in the response</span></span>
 
<span data-ttu-id="ea7f4-160">A formata??o de propriedades como `size` e `color` conter? valores `null` na resposta quando valores diferentes existirem no intervalo especificado.</span><span class="sxs-lookup"><span data-stu-id="ea7f4-160">Formatting properties such as `size` and `color` will contain `null` values in the response when different values exist in the specified range.</span></span> <span data-ttu-id="ea7f4-161">Por exemplo, se voc? recuperar um intervalo e carregar sua propriedade `format.font.color`:</span><span class="sxs-lookup"><span data-stu-id="ea7f4-161">For example, if you retrieve a range and load its `format.font.color` property:</span></span>
 
* <span data-ttu-id="ea7f4-162">Se todas as c?lulas no intervalo tiverem a mesma cor de fonte, `range.format.font.color` especificar? essa cor.</span><span class="sxs-lookup"><span data-stu-id="ea7f4-162">If all cells in the range have the same font color, `range.format.font.color` specifies that color.</span></span>
* <span data-ttu-id="ea7f4-163">Se houver v?rias cores de fonte dentro do intervalo, `range.format.font.color` ser? `null`.</span><span class="sxs-lookup"><span data-stu-id="ea7f4-163">If multiple font colors are present within the range, `range.format.font.color` is `null`.</span></span>
 
### <a name="blank-input-for-a-property"></a><span data-ttu-id="ea7f4-164">Entrada em branco para uma propriedade</span><span class="sxs-lookup"><span data-stu-id="ea7f4-164">Blank input for a property</span></span>
 
<span data-ttu-id="ea7f4-p118">Quando voc? especificar um valor em branco para uma propriedade (isto ?, duas aspas sem espa?o entre elas `''`), ele ser? interpretado como uma instru??o para limpar ou redefinir a propriedade. Por exemplo:</span><span class="sxs-lookup"><span data-stu-id="ea7f4-p118">When you specify a blank value for a property (i.e., two quotation marks with no space in-between `''`), it will be interpreted as an instruction to clear or reset the property. For example:</span></span>
 
* <span data-ttu-id="ea7f4-167">Se voc? especificar um valor em branco para a propriedade `values` de um intervalo, o conte?do do intervalo ser? apagado.</span><span class="sxs-lookup"><span data-stu-id="ea7f4-167">If you specify a blank value for the `values` property of a range, the content of the range is cleared.</span></span>
 
* <span data-ttu-id="ea7f4-168">Se voc? especificar um valor em branco para a propriedade `numberFormat`, o formato de n?mero ser? redefinido para `General`.</span><span class="sxs-lookup"><span data-stu-id="ea7f4-168">If you specify a blank value for the `numberFormat` property, the number format is reset to `General`.</span></span>
 
* <span data-ttu-id="ea7f4-169">Se voc? especificar um valor em branco para a propriedade `formula` e a propriedade `formulaLocale`, os valores de f?rmula ser?o apagados.</span><span class="sxs-lookup"><span data-stu-id="ea7f4-169">If you specify a blank value for the `formula` property and `formulaLocale` property, the formula values are cleared.</span></span>
 
### <a name="blank-property-values-in-the-response"></a><span data-ttu-id="ea7f4-170">Valores da propriedade em branco na resposta</span><span class="sxs-lookup"><span data-stu-id="ea7f4-170">Blank property values in the response</span></span>
 
<span data-ttu-id="ea7f4-171">Para opera??es de leitura, um valor de propriedade em branco na resposta (isto ?, duas aspas sem espa?o entre elas `''`) indica que a c?lula n?o cont?m dados nem valor.</span><span class="sxs-lookup"><span data-stu-id="ea7f4-171">For read operations, a blank property value in the response (i.e., two quotation marks with no space in-between `''`) indicates that cell contains no data or value.</span></span> <span data-ttu-id="ea7f4-172">No primeiro exemplo abaixo, a primeira e a ?ltima c?lula no intervalo n?o cont?m dados.</span><span class="sxs-lookup"><span data-stu-id="ea7f4-172">In the first example below, the first and last cell in the range contain no data.</span></span> <span data-ttu-id="ea7f4-173">No segundo exemplo, as primeiras duas c?lulas no intervalo n?o cont?m uma f?rmula.</span><span class="sxs-lookup"><span data-stu-id="ea7f4-173">In the second example, the first two cells in the range do not contain a formula.</span></span>
 
```js
range.values = [['', 'some', 'data', 'in', 'other', 'cells', '']];
```
 
```js
range.formula = [['', '', '=Rand()']];
```
 
## <a name="read-or-write-to-an-unbounded-range"></a><span data-ttu-id="ea7f4-174">Ler ou gravar em um intervalo n?o limitado</span><span class="sxs-lookup"><span data-stu-id="ea7f4-174">Read or write to an unbounded range</span></span>
 
### <a name="read-an-unbounded-range"></a><span data-ttu-id="ea7f4-175">Ler um intervalo n?o limitado</span><span class="sxs-lookup"><span data-stu-id="ea7f4-175">Read an unbounded range</span></span>
 
<span data-ttu-id="ea7f4-p120">Um endere?o de intervalo n?o limitado ? um endere?o de intervalo que especifica colunas ou linhas inteiras. Por exemplo:</span><span class="sxs-lookup"><span data-stu-id="ea7f4-p120">An unbounded range address is a range address that specifies either entire column(s) or entire row(s). For example:</span></span>
 
* <span data-ttu-id="ea7f4-178">Endere?os de intervalo composto por colunas inteiras:</span><span class="sxs-lookup"><span data-stu-id="ea7f4-178">Range addresses comprised of entire column(s):</span></span><ul><li>`C:C`</li><li>`A:F`</li></ul>
* <span data-ttu-id="ea7f4-179">Endere?os de intervalo composto por linhas inteiras:</span><span class="sxs-lookup"><span data-stu-id="ea7f4-179">Range addresses comprised of entire row(s):</span></span><ul><li>`2:2`</li><li>`1:4`</li></ul>
 
<span data-ttu-id="ea7f4-180">Quando uma API faz uma solicita??o para recuperar um intervalo n?o limitado (por exemplo, `getRange('C:C')`), a resposta conter? valores `null` para as propriedades no n?vel de c?lula, como `values`, `text`, `numberFormat` e `formula`.</span><span class="sxs-lookup"><span data-stu-id="ea7f4-180">When the API makes a request to retrieve an unbounded range (for example, `getRange('C:C')`), the response will contain `null` values for cell-level properties such as `values`, `text`, `numberFormat`, and `formula`.</span></span> <span data-ttu-id="ea7f4-181">Outras propriedades do intervalo, como `address` e `cellCount`, conter?o valores v?lidos para o intervalo n?o limitado.</span><span class="sxs-lookup"><span data-stu-id="ea7f4-181">Other properties of the range, such as `address` and `cellCount`, will contain valid values for the unbounded range.</span></span>
 
### <a name="write-to-an-unbounded-range"></a><span data-ttu-id="ea7f4-182">Gravar em um intervalo n?o limitado</span><span class="sxs-lookup"><span data-stu-id="ea7f4-182">Write to an unbounded range</span></span>
 
<span data-ttu-id="ea7f4-183">N?o ? poss?vel definir propriedades no n?vel de c?lula, como `values`, `numberFormat` e `formula`, no intervalo n?o limitado, pois a solicita??o de entrada ? muito grande.</span><span class="sxs-lookup"><span data-stu-id="ea7f4-183">You cannot set cell-level properties such as `values`, `numberFormat`, and `formula` on unbounded range because the input request is too large.</span></span> <span data-ttu-id="ea7f4-184">Por exemplo, o trecho de c?digo a seguir n?o ? v?lida porque ele tenta especificar `values` para um intervalo n?o limitado.</span><span class="sxs-lookup"><span data-stu-id="ea7f4-184">For example, the following code snippet is not valid because it attempts to specify `values` for an unbounded range.</span></span> <span data-ttu-id="ea7f4-185">A API retornar? um erro se voc? tentar definir as propriedades no n?vel de c?lula para um intervalo n?o limitado.</span><span class="sxs-lookup"><span data-stu-id="ea7f4-185">The API will return an error if you attempt to set cell-level properties for an unbounded range.</span></span>
 
```js
const range = context.workbook.worksheets.getActiveWorksheet().getRange('A:B');
range.values = 'Due Date';
```
 
## <a name="read-or-write-to-a-large-range"></a><span data-ttu-id="ea7f4-186">Ler ou gravar em um intervalo grande</span><span class="sxs-lookup"><span data-stu-id="ea7f4-186">Read or write to a large range</span></span>
 
<span data-ttu-id="ea7f4-187">Se um intervalo contiver um grande n?mero de c?lulas, valores, formatos de n?mero e/ou f?rmulas, talvez n?o seja poss?vel executar opera??es de API nesse intervalo.</span><span class="sxs-lookup"><span data-stu-id="ea7f4-187">If a range contains a large number of cells, values, number formats, and/or formulas, it may not be possible to run API operations on that range.</span></span> <span data-ttu-id="ea7f4-188">A API sempre far? a melhor tentativa de executar a opera??o solicitada em um intervalo (isto ?, para recuperar ou gravar os dados especificados), mas tentar executar opera??es de leitura ou grava??o para um intervalo grande pode resultar em um erro de API devido ? utiliza??o excessiva de recursos.</span><span class="sxs-lookup"><span data-stu-id="ea7f4-188">The API will always make a best attempt to run the requested operation on a range (i.e., to retrieve or write the specified data), but attempting to perform read or write operations for a large range may result in an API error due to excessive resource utilization.</span></span> <span data-ttu-id="ea7f4-189">Para evitar tais erros, ? recomend?vel executar opera??es de leitura ou grava??o separadas para subconjuntos menores de um intervalo grande, em vez de tentar executar uma ?nica opera??o de leitura ou grava??o em um intervalo grande.</span><span class="sxs-lookup"><span data-stu-id="ea7f4-189">To avoid such errors, we recommend that you run separate read or write operations for smaller subsets of a large range, instead of attempting to run a single read or write operation on a large range.</span></span>
 
## <a name="update-all-cells-in-a-range"></a><span data-ttu-id="ea7f4-190">Atualizar todas as c?lulas em um intervalo</span><span class="sxs-lookup"><span data-stu-id="ea7f4-190">Update all cells in a range</span></span>
 
<span data-ttu-id="ea7f4-191">Para aplicar a mesma atualiza??o a todas as c?lulas em um intervalo, (por exemplo, para popular todas as c?lulas com o mesmo valor, definir o mesmo formato de n?mero ou popular todas as c?lulas com a mesma f?rmula), defina a propriedade correspondente no objeto **range** para o valor (?nico) desejado.</span><span class="sxs-lookup"><span data-stu-id="ea7f4-191">To apply the same update to all cells in a range, (for example, to populate all cells with the same value, set the same number format, or populate all cells with the same formula), set the corresponding property on the **range** object to the desired (single) value.</span></span>
 
<span data-ttu-id="ea7f4-192">O exemplo a seguir obt?m um intervalo que cont?m 20 c?lulas e, em seguida, define o formato de n?mero e popula todas as c?lulas do intervalo com o valor **11/3/2015**.</span><span class="sxs-lookup"><span data-stu-id="ea7f4-192">The following example gets a range that contains 20 cells, and then sets the number format and populates all cells in the range with the value **3/11/2015**.</span></span>
 
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
 
## <a name="error-messages"></a><span data-ttu-id="ea7f4-193">Mensagens de erro</span><span class="sxs-lookup"><span data-stu-id="ea7f4-193">Error messages</span></span>
 
<span data-ttu-id="ea7f4-194">Quando ocorrer um erro de API, a API retornar? um objeto **error** que cont?m um c?digo e uma mensagem.</span><span class="sxs-lookup"><span data-stu-id="ea7f4-194">When an API error occurs, the API will return an **error** object that contains a code and a message.</span></span> <span data-ttu-id="ea7f4-195">A tabela a seguir define uma lista de erros que a API pode retornar.</span><span class="sxs-lookup"><span data-stu-id="ea7f4-195">The following table defines a list of errors that the API may return.</span></span>
 
|<span data-ttu-id="ea7f4-196">error.code</span><span class="sxs-lookup"><span data-stu-id="ea7f4-196">error.code</span></span> | <span data-ttu-id="ea7f4-197">error.message</span><span class="sxs-lookup"><span data-stu-id="ea7f4-197">error.message</span></span> |
|:----------|:--------------|
|<span data-ttu-id="ea7f4-198">InvalidArgument</span><span class="sxs-lookup"><span data-stu-id="ea7f4-198">InvalidArgument</span></span> |<span data-ttu-id="ea7f4-199">O argumento ? inv?lido, est? ausente ou tem um formato incorreto.</span><span class="sxs-lookup"><span data-stu-id="ea7f4-199">The argument is invalid or missing or has an incorrect format.</span></span>|
|<span data-ttu-id="ea7f4-200">InvalidRequest</span><span class="sxs-lookup"><span data-stu-id="ea7f4-200">InvalidRequest</span></span>  |<span data-ttu-id="ea7f4-201">N?o ? poss?vel processar a solicita??o.</span><span class="sxs-lookup"><span data-stu-id="ea7f4-201">Cannot process the request.</span></span>|
|<span data-ttu-id="ea7f4-202">InvalidReference</span><span class="sxs-lookup"><span data-stu-id="ea7f4-202">InvalidReference</span></span>|<span data-ttu-id="ea7f4-203">Esta refer?ncia n?o ? v?lida para a opera??o atual.</span><span class="sxs-lookup"><span data-stu-id="ea7f4-203">This reference is not valid for the current operation.</span></span>|
|<span data-ttu-id="ea7f4-204">InvalidBinding</span><span class="sxs-lookup"><span data-stu-id="ea7f4-204">InvalidBinding</span></span>  |<span data-ttu-id="ea7f4-205">Esta associa??o de objetos n?o ? mais v?lida devido ?s atualiza??es anteriores.</span><span class="sxs-lookup"><span data-stu-id="ea7f4-205">This object binding is no longer valid due to previous updates.</span></span>|
|<span data-ttu-id="ea7f4-206">InvalidSelection</span><span class="sxs-lookup"><span data-stu-id="ea7f4-206">InvalidSelection</span></span>|<span data-ttu-id="ea7f4-207">A sele??o atual ? inv?lida para esta opera??o.</span><span class="sxs-lookup"><span data-stu-id="ea7f4-207">The current selection is invalid for this operation.</span></span>|
|<span data-ttu-id="ea7f4-208">Unauthenticated</span><span class="sxs-lookup"><span data-stu-id="ea7f4-208">Unauthenticated</span></span> |<span data-ttu-id="ea7f4-209">Informa??es de autentica??o necess?rias est?o ausentes ou inv?lidas.</span><span class="sxs-lookup"><span data-stu-id="ea7f4-209">Required authentication information is either missing or invalid.</span></span>|
|<span data-ttu-id="ea7f4-210">AccessDenied</span><span class="sxs-lookup"><span data-stu-id="ea7f4-210">AccessDenied</span></span> |<span data-ttu-id="ea7f4-211">Voc? n?o pode realizar a opera??o solicitada.</span><span class="sxs-lookup"><span data-stu-id="ea7f4-211">You cannot perform the requested operation.</span></span>|
|<span data-ttu-id="ea7f4-212">ItemNotFound</span><span class="sxs-lookup"><span data-stu-id="ea7f4-212">ItemNotFound</span></span> |<span data-ttu-id="ea7f4-213">O recurso solicitado n?o existe.</span><span class="sxs-lookup"><span data-stu-id="ea7f4-213">The requested resource doesn't exist.</span></span>|
|<span data-ttu-id="ea7f4-214">ActivityLimitReached</span><span class="sxs-lookup"><span data-stu-id="ea7f4-214">ActivityLimitReached</span></span>|<span data-ttu-id="ea7f4-215">O limite de atividades foi alcan?ado.</span><span class="sxs-lookup"><span data-stu-id="ea7f4-215">Activity limit has been reached.</span></span>|
|<span data-ttu-id="ea7f4-216">GeneralException</span><span class="sxs-lookup"><span data-stu-id="ea7f4-216">GeneralException</span></span>|<span data-ttu-id="ea7f4-217">Ocorreu um erro interno ao processar a solicita??o.</span><span class="sxs-lookup"><span data-stu-id="ea7f4-217">There was an internal error while processing the request.</span></span>|
|<span data-ttu-id="ea7f4-218">NotImplemented</span><span class="sxs-lookup"><span data-stu-id="ea7f4-218">NotImplemented</span></span>  |<span data-ttu-id="ea7f4-219">O recurso solicitado n?o foi implementado.</span><span class="sxs-lookup"><span data-stu-id="ea7f4-219">The requested feature isn't implemented.</span></span>|
|<span data-ttu-id="ea7f4-220">ServiceNotAvailable</span><span class="sxs-lookup"><span data-stu-id="ea7f4-220">ServiceNotAvailable</span></span>|<span data-ttu-id="ea7f4-221">O servi?o n?o est? dispon?vel.</span><span class="sxs-lookup"><span data-stu-id="ea7f4-221">The service is unavailable.</span></span>|
|<span data-ttu-id="ea7f4-222">Conflito</span><span class="sxs-lookup"><span data-stu-id="ea7f4-222">Conflict</span></span>              |<span data-ttu-id="ea7f4-223">A solicita??o n?o p?de ser processada devido a um conflito.</span><span class="sxs-lookup"><span data-stu-id="ea7f4-223">Request could not be processed because of a conflict.</span></span>|
|<span data-ttu-id="ea7f4-224">ItemAlreadyExists</span><span class="sxs-lookup"><span data-stu-id="ea7f4-224">ItemAlreadyExists</span></span>|<span data-ttu-id="ea7f4-225">O recurso que est? sendo criado j? existe.</span><span class="sxs-lookup"><span data-stu-id="ea7f4-225">The resource being created already exists.</span></span>|
|<span data-ttu-id="ea7f4-226">UnsupportedOperation</span><span class="sxs-lookup"><span data-stu-id="ea7f4-226">UnsupportedOperation</span></span>|<span data-ttu-id="ea7f4-227">N?o h? suporte para a opera??o que est? sendo tentada.</span><span class="sxs-lookup"><span data-stu-id="ea7f4-227">The operation being attempted is not supported.</span></span>|
|<span data-ttu-id="ea7f4-228">RequestAborted</span><span class="sxs-lookup"><span data-stu-id="ea7f4-228">RequestAborted</span></span>|<span data-ttu-id="ea7f4-229">A solicita??o foi anulada durante o tempo de execu??o.</span><span class="sxs-lookup"><span data-stu-id="ea7f4-229">The request was aborted during run time.</span></span>|
|<span data-ttu-id="ea7f4-230">ApiNotAvailable</span><span class="sxs-lookup"><span data-stu-id="ea7f4-230">ApiNotAvailable</span></span>|<span data-ttu-id="ea7f4-231">A API solicitada n?o est? dispon?vel.</span><span class="sxs-lookup"><span data-stu-id="ea7f4-231">The requested API is not available.</span></span>|
|<span data-ttu-id="ea7f4-232">InsertDeleteConflict</span><span class="sxs-lookup"><span data-stu-id="ea7f4-232">InsertDeleteConflict</span></span>|<span data-ttu-id="ea7f4-233">A tentativa de opera??o de exclus?o ou inser??o resultou em um conflito.</span><span class="sxs-lookup"><span data-stu-id="ea7f4-233">The insert or delete operation attempted resulted in a conflict.</span></span>|
|<span data-ttu-id="ea7f4-234">InvalidOperation</span><span class="sxs-lookup"><span data-stu-id="ea7f4-234">InvalidOperation</span></span>|<span data-ttu-id="ea7f4-235">A tentativa de opera??o ? inv?lida no objeto.</span><span class="sxs-lookup"><span data-stu-id="ea7f4-235">The operation attempted is invalid on the object.</span></span>|
 
## <a name="see-also"></a><span data-ttu-id="ea7f4-236">Veja tamb?m</span><span class="sxs-lookup"><span data-stu-id="ea7f4-236">See also</span></span>
 
* [<span data-ttu-id="ea7f4-237">Introdu??o aos suplementos do Excel</span><span class="sxs-lookup"><span data-stu-id="ea7f4-237">Get started with Excel add-ins</span></span>](excel-add-ins-get-started-overview.md)
* [<span data-ttu-id="ea7f4-238">Exemplos de c?digo de suplementos do Excel</span><span class="sxs-lookup"><span data-stu-id="ea7f4-238">Excel add-ins code samples</span></span>](http://dev.office.com/code-samples#?filters=excel,office%20add-ins)
* [<span data-ttu-id="ea7f4-239">Otimiza??o de desempenho da API JavaScript do Excel</span><span class="sxs-lookup"><span data-stu-id="ea7f4-239">Excel JavaScript API performance optimization</span></span>](https://dev.office.com/reference/add-ins/excel/performance.md)
* [<span data-ttu-id="ea7f4-240">Refer?ncia da API JavaScript do Excel</span><span class="sxs-lookup"><span data-stu-id="ea7f4-240">Excel JavaScript API reference</span></span>](https://dev.office.com/reference/add-ins/excel/excel-add-ins-reference-overview)
