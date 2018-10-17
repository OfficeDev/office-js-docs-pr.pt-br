---
title: Conceitos fundamentais de programação com a API JavaScript do Excel
description: Usar a API JavaScript do Excel para criar suplementos para o Excel.
ms.date: 10/03/2018
ms.openlocfilehash: f93ec7b5e34f90f2d61f29d861b7e0c19f66f6e3
ms.sourcegitcommit: c53f05bbd4abdfe1ee2e42fdd4f82b318b363ad7
ms.translationtype: HT
ms.contentlocale: pt-BR
ms.lasthandoff: 10/12/2018
ms.locfileid: "25505983"
---
# <a name="fundamental-programming-concepts-with-the-excel-javascript-api"></a><span data-ttu-id="e729f-103">Conceitos fundamentais de programação com a API JavaScript do Excel</span><span class="sxs-lookup"><span data-stu-id="e729f-103">Fundamental programming concepts with the Excel JavaScript API</span></span>
 
<span data-ttu-id="e729f-p101">Este artigo descreve como usar a [API JavaScript do Excel](https://docs.microsoft.com/office/dev/add-ins/reference/overview/excel-add-ins-reference-overview?view=office-js) para criar suplementos para o Excel 2016 ou posterior. Ele apresenta os principais conceitos fundamentais para o uso da API e fornece orientação para a realização de tarefas específicas, como fazer a leitura ou gravação em um intervalo grande, atualizar todas as células de um intervalo e muito mais.</span><span class="sxs-lookup"><span data-stu-id="e729f-p101">This article describes how to use the [Excel JavaScript API](https://docs.microsoft.com/office/dev/add-ins/reference/overview/excel-add-ins-reference-overview?view=office-js) to build add-ins for Excel 2016 or later. It introduces core concepts that are fundamental to using the API and provides guidance for performing specific tasks such as reading or writing to a large range, updating all cells in range, and more.</span></span>

## <a name="asynchronous-nature-of-excel-apis"></a><span data-ttu-id="e729f-106">Natureza assíncrona das APIs do Excel</span><span class="sxs-lookup"><span data-stu-id="e729f-106">Asynchronous nature of Excel APIs</span></span>

<span data-ttu-id="e729f-p102">Os suplementos do Excel baseados na Web são executados dentro de um contêiner de navegador incorporado ao aplicativo do Office em plataformas baseadas em área de trabalho, como o Office para o Windows, e é executado dentro de um iFrame HTML no Office Online. Não é viável habilitar a API Office.js para interagir de maneira síncrona com o host do Excel em todas as plataformas com suporte devido a considerações de desempenho. Portanto, a chamada **sync()** de API no Office.js retorna uma [promessa](https://developer.mozilla.org/docs/Web/JavaScript/Reference/Global_Objects/Promise) que é resolvida quando o aplicativo do Excel conclui as ações de leitura ou gravação solicitadas. Além disso, você pode adicionar várias ações, como definir propriedades ou métodos, a uma fila e executá-las como um lote de comandos com uma única chamada **sync()**, ao invés de enviar uma solicitação separada para cada ação. As seções a seguir descrevem como realizar essa tarefa usando as APIs **Excel.run()** e **sync()**.</span><span class="sxs-lookup"><span data-stu-id="e729f-p102">The web-based Excel add-ins run inside a browser container that is embedded within the Office application on desktop-based platforms such as Office for Windows and runs inside an HTML iFrame in Office Online. Enabling the Office.js API to interact synchronously with the Excel host across all supported platforms is not feasible due to performance considerations. Therefore, the **sync()** API call in Office.js returns a [promise](https://developer.mozilla.org/docs/Web/JavaScript/Reference/Global_Objects/Promise) that is resolved when the Excel application completes the requested read or write actions. Also, you can queue up multiple actions, such as setting properties or invoking methods, and run them as a batch of commands with a single call to **sync()**, rather than sending a separate request for each action. The following sections describe how to accomplish this using the **Excel.run()** and **sync()** APIs.</span></span>
 
## <a name="excelrun"></a><span data-ttu-id="e729f-112">Excel.run</span><span class="sxs-lookup"><span data-stu-id="e729f-112">Excel.run</span></span>
 
<span data-ttu-id="e729f-p103">**Excel.Run** executa uma função onde você pode especificar as ações a serem executadas em relação ao modelo de objeto do Excel. **Excel.Run** cria automaticamente um contexto de solicitação que você pode usar para interagir com objetos do Excel. Quando **Excel.run** é concluída, uma promessa é resolvida e todos os objetos alocados durante o tempo de execução são automaticamente liberados.</span><span class="sxs-lookup"><span data-stu-id="e729f-p103">**Excel.run** executes a function where you specify the actions to perform against the Excel object model. **Excel.run** automatically creates a request context that you can use to interact with Excel objects. When **Excel.run** completes, a promise is resolved, and any objects that were allocated at runtime are automatically released.</span></span>
 
<span data-ttu-id="e729f-p104">O exemplo a seguir mostra como usar **Excel.run**. A instrução catch captura e registra os erros que ocorrem em **Excel.run**.</span><span class="sxs-lookup"><span data-stu-id="e729f-p104">The following example shows how to use **Excel.run**. The catch statement catches and logs errors that occur within the **Excel.run**.</span></span>
 
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

## <a name="request-context"></a><span data-ttu-id="e729f-118">Contexto de solicitação</span><span class="sxs-lookup"><span data-stu-id="e729f-118">Request context</span></span>
 
<span data-ttu-id="e729f-p105">O Excel e o seu suplemento são executados em dois processos diferentes. Como eles usam diferentes ambientes de tempo de execução, os suplementos do Excel exigem um objeto **RequestContext** para conectar o suplemento aos objetos no Excel, como planilhas, intervalos, gráficos e tabelas.</span><span class="sxs-lookup"><span data-stu-id="e729f-p105">Excel and your add-in run in two different processes. Since they use different runtime environments, Excel add-ins require a **RequestContext** object in order to connect your add-in to objects in Excel such as worksheets, ranges, charts, and tables.</span></span>
 
## <a name="proxy-objects"></a><span data-ttu-id="e729f-121">Objetos proxy</span><span class="sxs-lookup"><span data-stu-id="e729f-121">Proxy objects</span></span>
 
<span data-ttu-id="e729f-p106">Os objetos do Excel JavaScript que você declara e usa em um suplemento são objetos proxy. Qualquer método invocado ou propriedade definida ou carregada por você nos objetos proxy simplesmente são adicionadas a uma fila de comandos pendentes. Quando você chama o método **sync()** no contexto da solicitação (por exemplo, `context.sync()`), os comandos na fila são enviados para o Excel e executados. A API JavaScript do Excel é fundamentalmente centrada em lotes. Você colocar quantas alterações desejar na fila no contexto da solicitação e, então, chamar o método **sync()** para executar o lote de comandos.</span><span class="sxs-lookup"><span data-stu-id="e729f-p106">The Excel JavaScript objects that you declare and use in an add-in are proxy objects. Any methods that you invoke or properties that you set or load on proxy objects are simply added to a queue of pending commands. When you call the **sync()** method on the request context (for example, `context.sync()`), the queued commands are dispatched to Excel and run. The Excel JavaScript API is fundamentally batch-centric. You can queue up as many changes as you wish on the request context, and then call the **sync()** method to run the batch of queued commands.</span></span>
 
<span data-ttu-id="e729f-p107">Por exemplo, o trecho de código a seguir declara o objeto JavaScript **selectedRange** local para fazer referência a um intervalo selecionado no documento do Excel e, em seguida, define algumas propriedades nesse objeto. O objeto **selectedRange** é um objeto proxy, portanto, as propriedades definidas e o método invocado em nele não refletem no documento do Excel até que seu suplemento chame **context.sync()**.</span><span class="sxs-lookup"><span data-stu-id="e729f-p107">For example, the following code snippet declares the local JavaScript object **selectedRange** to reference a selected range in the Excel document, and then sets some properties on that object. The **selectedRange** object is a proxy object, so the properties that are set and method that is invoked on that object will not be reflected in the Excel document until your add-in calls **context.sync()**.</span></span>
 
```js
const selectedRange = context.workbook.getSelectedRange();
selectedRange.format.fill.color = "#4472C4";
selectedRange.format.font.color = "white";
selectedRange.format.autofitColumns();
```
 
### <a name="sync"></a><span data-ttu-id="e729f-129">sync()</span><span class="sxs-lookup"><span data-stu-id="e729f-129">sync()</span></span>
 
<span data-ttu-id="e729f-p108">Chamar o método **sync()** no contexto da solicitação sincroniza o estado entre os objetos proxy e os objetos no documento do Excel. O método **sync()** executa os comandos na fila no contexto da solicitação e recupera os valores para todas as propriedades que devem ser carregadas nos objetos proxy. O método **sync()** é executado de forma assíncrona e retorna uma [promessa](https://developer.mozilla.org/docs/Web/JavaScript/Reference/Global_Objects/Promise), que é resolvida quando o método **sync()** é concluído.</span><span class="sxs-lookup"><span data-stu-id="e729f-p108">Calling the **sync()** method on the request context synchronizes the state between proxy objects and objects in the Excel document. The **sync()** method runs any commands that are queued on the request context and retrieves values for any properties that should be loaded on the proxy objects. The **sync()** method executes asynchronously and returns a [promise](https://developer.mozilla.org/docs/Web/JavaScript/Reference/Global_Objects/Promise), which is resolved when the **sync()** method completes.</span></span>
 
<span data-ttu-id="e729f-133">O exemplo a seguir mostra uma função de lote que define um objeto proxy JavaScript local (**selectedRange**), carrega uma propriedade desse objeto e, em seguida, usa o padrão JavaScript Promises para chamar **context.sync()** a fim de sincronizar o estado entre objetos proxy e objetos no documento do Excel.</span><span class="sxs-lookup"><span data-stu-id="e729f-133">The following example shows a batch function that defines a local JavaScript proxy object (**selectedRange**), loads a property of that object, and then uses the JavaScript Promises pattern to call **context.sync()** to synchronize the state between proxy objects and objects in the Excel document.</span></span>
 
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
 
<span data-ttu-id="e729f-134">No exemplo anterior, **selectedRange** está definido e sua propriedade **address** é carregada quando **context.sync()** é chamado.</span><span class="sxs-lookup"><span data-stu-id="e729f-134">In the previous example, **selectedRange** is set and its **address** property is loaded when **context.sync()** is called.</span></span>
 
<span data-ttu-id="e729f-p109">Como **sync()** é uma operação assíncrona que retorna uma promessa, você sempre deve **retornar** a promessa (em JavaScript). Isso garante que a operação **sync()** seja concluída antes do script continuar a ser executado. Para obter mais informações sobre como otimizar o desempenho com **sync()**, consulte [Otimização do desempenho da API JavaScript do Excel](https://docs.microsoft.com/office/dev/add-ins/excel/performance).</span><span class="sxs-lookup"><span data-stu-id="e729f-p109">Because **sync()** is an asynchronous operation that returns a promise, you should always **return** the promise (in JavaScript). Doing so ensures that the **sync()** operation completes before the script continues to run. For more information about optimizing performance with **sync()**, see [Excel JavaScript API performance optimization](https://docs.microsoft.com/office/dev/add-ins/excel/performance).</span></span>
 
### <a name="load"></a><span data-ttu-id="e729f-138">load()</span><span class="sxs-lookup"><span data-stu-id="e729f-138">load()</span></span>
 
<span data-ttu-id="e729f-p110">Antes de poder ler as propriedades de um objeto proxy, você deve carregar explicitamente as propriedades para preencher o objeto com dados de um documento do Excel e, em seguida, chamar **context.sync()**. Por exemplo, se você criar um objeto proxy para fazer referência a um intervalo selecionado e quiser ler a propriedade **address** desse intervalo, você primeiro precisará carregar a propriedade **address**. Para solicitar que uma propriedade de um objeto proxy seja carregada, chame o método **load()** no objeto e especifique as propriedades que devem ser carregadas.</span><span class="sxs-lookup"><span data-stu-id="e729f-p110">Before you can read the properties of a proxy object, you must explicitly load the properties to populate the proxy object with data from the Excel document, and then call **context.sync()**. For example, if you create a proxy object to reference a selected range, and then want to read the selected range's **address** property, you need to load the **address** property before you can read it. To request properties of a proxy object be loaded, call the **load()** method on the object and specify the properties to load.</span></span> 

> [!NOTE]
> <span data-ttu-id="e729f-p111">Se você estiver apenas chamando métodos ou definindo propriedades em um objeto proxy, você não precisa chamar o método **load()**. O método **load()** só é necessário quando você deseja ler as propriedades em um objeto proxy.</span><span class="sxs-lookup"><span data-stu-id="e729f-p111">If you are only calling methods or setting properties on a proxy object, you do not need to call the **load()** method. The **load()** method is only required when you want to read properties on a proxy object.</span></span>
 
<span data-ttu-id="e729f-p112">Assim como as solicitações para definir propriedades ou invocar métodos em objetos proxy, as solicitações para carregar propriedades em objetos proxy são adicionadas à fila de comandos pendentes no contexto da solicitação, sendo executadas na próxima vez que você chamar o método **sync()** . É possível colocar quantas chamadas de **load()** forem necessárias na fila no contexto da solicitação.</span><span class="sxs-lookup"><span data-stu-id="e729f-p112">Just like requests to set properties or invoke methods on proxy objects, requests to load properties on proxy objects get added to the queue of pending commands on the request context, which will run the next time you call the **sync()** method. You can queue up as many **load()** calls on the request context as necessary.</span></span>
 
<span data-ttu-id="e729f-146">No exemplo a seguir, somente propriedades específicas do intervalo são carregadas.</span><span class="sxs-lookup"><span data-stu-id="e729f-146">In the following example, only specific properties of the range are loaded.</span></span>
 
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
 
<span data-ttu-id="e729f-147">No exemplo anterior, como `format/font` não é especificado na chamada a **myRange.load()**, a propriedade `format.font.color` não pode ser lida.</span><span class="sxs-lookup"><span data-stu-id="e729f-147">In the previous example, because `format/font` is not specified in the call to **myRange.load()**, the `format.font.color` property cannot be read.</span></span>

<span data-ttu-id="e729f-148">Para otimizar o desempenho, você deve especificar explicitamente as propriedades e relações a serem carregadas ao usar o método **load()** em um objeto, conforme abordado em [Otimizações de desempenho da API JavaScript do Excel](performance.md).</span><span class="sxs-lookup"><span data-stu-id="e729f-148">To optimize performance, you should explicitly specify the properties and relationships to load when using the **load()** method on an object, as covered in [Excel JavaScript API performance optimizations](performance.md).</span></span> <span data-ttu-id="e729f-149">Para obter mais informações sobre o método **Load** , consulte [Conceitos de programação avançados com a API JavaScript do Excel](excel-add-ins-advanced-concepts.md).</span><span class="sxs-lookup"><span data-stu-id="e729f-149">For more information about the **load()** method, see [Advanced programming concepts with the Excel JavaScript API](excel-add-ins-advanced-concepts.md).</span></span>

## <a name="null-or-blank-property-values"></a><span data-ttu-id="e729f-150">Valores de propriedade null ou blank</span><span class="sxs-lookup"><span data-stu-id="e729f-150">null or blank property values</span></span>
 
### <a name="null-input-in-2-d-array"></a><span data-ttu-id="e729f-151">entrada nula em uma matriz 2D</span><span class="sxs-lookup"><span data-stu-id="e729f-151">null input in 2-D Array</span></span>
 
<span data-ttu-id="e729f-p114">No Excel, um intervalo é representado por uma matriz 2-D, onde a primeira dimensão é formada por linhas e a segunda por colunas. Para definir valores, formatos de número ou fórmulas para células específicas dentro de um intervalo, especifique os valores, formatos de número ou fórmulas para essas células na matriz 2D e especifique `null` para todas as outras células.</span><span class="sxs-lookup"><span data-stu-id="e729f-p114">In Excel, a range is represented by a 2-D array, where the first dimension is rows and the second dimension is columns. To set values, number format, or formula for only specific cells within a range, specify the values, number format, or formula for those cells in the 2-D array, and specify `null` for all other cells in the 2-D array.</span></span>
 
<span data-ttu-id="e729f-p115">Por exemplo, para atualizar o formato de número de apenas uma célula dentro de um intervalo de atualização e manter o formato existente para todas as outras células, especifique o novo formato para a célula que deseja atualizar e especifique `null` para todas as outras células. Os trechos de código a seguir definem um novo formato de número para a quarta célula do intervalo e mantém o formato inalterado para as três primeiras células no intervalo.</span><span class="sxs-lookup"><span data-stu-id="e729f-p115">For example, to update the number format for only one cell within a range, and retain the existing number format for all other cells in the range, specify the new number format for the cell to update, and specify `null` for all other cells. The following code snippet sets a new number format for the fourth cell in the range, and leaves the number format unchanged for the first three cells in the range.</span></span>
 
```js
range.values = [['Eurasia', '29.96', '0.25', '15-Feb' ]];
range.numberFormat = [[null, null, null, 'm/d/yyyy;@']];
```
 
### <a name="null-input-for-a-property"></a><span data-ttu-id="e729f-156">entrada nula para uma propriedade</span><span class="sxs-lookup"><span data-stu-id="e729f-156">null input for a property</span></span>
 
<span data-ttu-id="e729f-p116">`null` não é uma entrada válida para uma única propriedade. Por exemplo, o snippet de código a seguir não é válido, pois a propriedade **values** do intervalo não pode ser definida como `null`.</span><span class="sxs-lookup"><span data-stu-id="e729f-p116">`null` is not a valid input for single property. For example, the following code snippet is not valid, as the **values** property of the range cannot be set to `null`.</span></span>
 
```js
range.values = null;
```
 
<span data-ttu-id="e729f-159">Da mesma forma, o snippet de código a seguir não é válido, pois `null` não é um valor válido para a propriedade **color**.</span><span class="sxs-lookup"><span data-stu-id="e729f-159">Likewise, the following code snippet is not valid, as `null` is not a valid value for the **color** property.</span></span>
 
```js
range.format.fill.color =  null;
```
 
### <a name="null-property-values-in-the-response"></a><span data-ttu-id="e729f-160">Valores nulos para propriedades na resposta</span><span class="sxs-lookup"><span data-stu-id="e729f-160">null property values in the response</span></span>
 
<span data-ttu-id="e729f-p117">Propriedades de formatação, como `size` e `color` contêm valores `null` na resposta quando há valores diferentes no intervalo especificado. Por exemplo, se você recuperar um intervalo e carregar a sua propriedade `format.font.color`:</span><span class="sxs-lookup"><span data-stu-id="e729f-p117">Formatting properties such as `size` and `color` will contain `null` values in the response when different values exist in the specified range. For example, if you retrieve a range and load its `format.font.color` property:</span></span>
 
* <span data-ttu-id="e729f-163">Se todas as células no intervalo tiverem a mesma cor de fonte, `range.format.font.color` especifica essa cor.</span><span class="sxs-lookup"><span data-stu-id="e729f-163">If all cells in the range have the same font color, `range.format.font.color` specifies that color.</span></span>
* <span data-ttu-id="e729f-164">Se houver várias cores de fonte dentro do intervalo, `range.format.font.color` será `null`.</span><span class="sxs-lookup"><span data-stu-id="e729f-164">If multiple font colors are present within the range, `range.format.font.color` is `null`.</span></span>
 
### <a name="blank-input-for-a-property"></a><span data-ttu-id="e729f-165">Entrada em branco para uma propriedade</span><span class="sxs-lookup"><span data-stu-id="e729f-165">Blank input for a property</span></span>
 
<span data-ttu-id="e729f-p118">Quando você especifica um valor em branco para uma propriedade (isto é, duas aspas sem espaço `''`), ele será interpretado como uma instrução para limpar ou redefinir a propriedade. Por exemplo:</span><span class="sxs-lookup"><span data-stu-id="e729f-p118">When you specify a blank value for a property (i.e., two quotation marks with no space in-between `''`), it will be interpreted as an instruction to clear or reset the property. For example:</span></span>
 
* <span data-ttu-id="e729f-168">Se você especificar um valor em branco para a propriedade `values` de um intervalo, o conteúdo do intervalo será apagado.</span><span class="sxs-lookup"><span data-stu-id="e729f-168">If you specify a blank value for the `values` property of a range, the content of the range is cleared.</span></span>
 
* <span data-ttu-id="e729f-169">Se você especificar um valor em branco para a propriedade `numberFormat`, o formato de número será redefinido para `General`.</span><span class="sxs-lookup"><span data-stu-id="e729f-169">If you specify a blank value for the `numberFormat` property, the number format is reset to `General`.</span></span>
 
* <span data-ttu-id="e729f-170">Se você especificar um valor em branco para as propriedades `formula` e `formulaLocale`, os valores de fórmula serão apagados.</span><span class="sxs-lookup"><span data-stu-id="e729f-170">If you specify a blank value for the `formula` property and `formulaLocale` property, the formula values are cleared.</span></span>
 
### <a name="blank-property-values-in-the-response"></a><span data-ttu-id="e729f-171">Valores de propriedade em branco na resposta</span><span class="sxs-lookup"><span data-stu-id="e729f-171">Blank property values in the response</span></span>
 
<span data-ttu-id="e729f-p119">Para operações de leitura, um valor de uma propriedade em branco na resposta (ou seja, duas aspas sem espaço `''`) indica que a célula não contém nenhum dado ou valor. No primeiro exemplo a seguir, a primeira e a última célula no intervalo não contêm nenhum dado. No segundo exemplo, as duas primeiras células no intervalo não contém uma fórmula.</span><span class="sxs-lookup"><span data-stu-id="e729f-p119">For read operations, a blank property value in the response (i.e., two quotation marks with no space in-between `''`) indicates that cell contains no data or value. In the first example below, the first and last cell in the range contain no data. In the second example, the first two cells in the range do not contain a formula.</span></span>
 
```js
range.values = [['', 'some', 'data', 'in', 'other', 'cells', '']];
```
 
```js
range.formula = [['', '', '=Rand()']];
```
 
## <a name="read-or-write-to-an-unbounded-range"></a><span data-ttu-id="e729f-175">Ler ou gravar em um intervalo não associado</span><span class="sxs-lookup"><span data-stu-id="e729f-175">Read or write to an unbounded range</span></span>
 
### <a name="read-an-unbounded-range"></a><span data-ttu-id="e729f-176">Ler um intervalo não associado</span><span class="sxs-lookup"><span data-stu-id="e729f-176">Read an unbounded range</span></span>
 
<span data-ttu-id="e729f-p120">Um endereço de intervalo não associado é um endereço de intervalo que especifica colunas ou linhas inteiras. Por exemplo:</span><span class="sxs-lookup"><span data-stu-id="e729f-p120">An unbounded range address is a range address that specifies either entire column(s) or entire row(s). For example:</span></span>
 
* <span data-ttu-id="e729f-179">Endereços de intervalo compostos por colunas inteiras:</span><span class="sxs-lookup"><span data-stu-id="e729f-179">Range addresses comprised of entire column(s):</span></span><ul><li>`C:C`</li><li>`A:F`</li></ul>
* <span data-ttu-id="e729f-180">Endereços de intervalo compostos por linhas inteiras:</span><span class="sxs-lookup"><span data-stu-id="e729f-180">Range addresses comprised of entire row(s):</span></span><ul><li>`2:2`</li><li>`1:4`</li></ul>
 
<span data-ttu-id="e729f-p121">Quando a API faz uma solicitação para recuperar um intervalo não associado (por exemplo, `getRange('C:C')`), a resposta contém valores `null` para propriedades de nível de célula, tais como `values`, `text`, `numberFormat`, e `formula`. Outras propriedades do intervalo, como `address` e `cellCount`, contêm valores válidos para o intervalo não associado.</span><span class="sxs-lookup"><span data-stu-id="e729f-p121">When the API makes a request to retrieve an unbounded range (for example, `getRange('C:C')`), the response will contain `null` values for cell-level properties such as `values`, `text`, `numberFormat`, and `formula`. Other properties of the range, such as `address` and `cellCount`, will contain valid values for the unbounded range.</span></span>
 
### <a name="write-to-an-unbounded-range"></a><span data-ttu-id="e729f-183">Gravar em um intervalo não associado</span><span class="sxs-lookup"><span data-stu-id="e729f-183">Write to an unbounded range</span></span>
 
<span data-ttu-id="e729f-p122">Você não pode definir propriedades em nível de célula, como `values`, `numberFormat`, e `formula`, em um intervalo não associado, pois a solicitação de entrada é muito grande. Por exemplo, o snippet de código a seguir não é válido pois tenta especificar `values` para um intervalo não associado. A API retornará um erro se você tentar definir propriedades de nível de célula para um intervalo não associado.</span><span class="sxs-lookup"><span data-stu-id="e729f-p122">You cannot set cell-level properties such as `values`, `numberFormat`, and `formula` on unbounded range because the input request is too large. For example, the following code snippet is not valid because it attempts to specify `values` for an unbounded range. The API will return an error if you attempt to set cell-level properties for an unbounded range.</span></span>
 
```js
const range = context.workbook.worksheets.getActiveWorksheet().getRange('A:B');
range.values = 'Due Date';
```
 
## <a name="read-or-write-to-a-large-range"></a><span data-ttu-id="e729f-187">Ler ou gravar em um intervalo grande</span><span class="sxs-lookup"><span data-stu-id="e729f-187">Read or write to a large range</span></span>
 
<span data-ttu-id="e729f-p123">Se um intervalo contiver um grande número de células, valores, formatos de número e/ou fórmulas, pode ser que não seja possível executar operações de API nesse intervalo. A API sempre fará a melhor tentativa para executar a operação solicitada em um intervalo (ou seja, recuperar ou gravar os dados especificados), mas a tentativa de executar operações de leitura ou gravação em um intervalo grande pode resultar em um erro de API devido à utilização excessiva de recursos. Para evitar esses erros, recomendamos que você execute operações de leitura ou gravação separadas para subconjuntos menores de um intervalo grande, em vez de tentar executar uma única operação em um intervalo grande.</span><span class="sxs-lookup"><span data-stu-id="e729f-p123">If a range contains a large number of cells, values, number formats, and/or formulas, it may not be possible to run API operations on that range. The API will always make a best attempt to run the requested operation on a range (i.e., to retrieve or write the specified data), but attempting to perform read or write operations for a large range may result in an API error due to excessive resource utilization. To avoid such errors, we recommend that you run separate read or write operations for smaller subsets of a large range, instead of attempting to run a single read or write operation on a large range.</span></span>
 
## <a name="update-all-cells-in-a-range"></a><span data-ttu-id="e729f-191">Atualizar todas as células em um intervalo</span><span class="sxs-lookup"><span data-stu-id="e729f-191">Update all cells in a range</span></span>
 
<span data-ttu-id="e729f-192">Para aplicar a mesma atualização a todas as células em um intervalo, (por exemplo, popular todas as células com o mesmo valor, definir o mesmo formato de número ou popular todas as células com a mesma fórmula), defina a propriedade correspondente no objeto **range** com o valor (único) desejado.</span><span class="sxs-lookup"><span data-stu-id="e729f-192">To apply the same update to all cells in a range, (for example, to populate all cells with the same value, set the same number format, or populate all cells with the same formula), set the corresponding property on the **range** object to the desired (single) value.</span></span>
 
<span data-ttu-id="e729f-193">O exemplo a seguir obtém um intervalo que contém 20 células e, em seguida, define o formato de número e popula todas as células do intervalo com o valor **11/3/2015**.</span><span class="sxs-lookup"><span data-stu-id="e729f-193">The following example gets a range that contains 20 cells, and then sets the number format and populates all cells in the range with the value **3/11/2015**.</span></span>
 
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
 
## <a name="error-messages"></a><span data-ttu-id="e729f-194">Mensagens de erro</span><span class="sxs-lookup"><span data-stu-id="e729f-194">Error messages</span></span>
 
<span data-ttu-id="e729f-p124">Quando ocorre um erro de API, a API retorna um objeto de **erro** que contém um código e uma mensagem. A tabela a seguir define uma lista de erros que a API pode retornar.</span><span class="sxs-lookup"><span data-stu-id="e729f-p124">When an API error occurs, the API will return an **error** object that contains a code and a message. The following table defines a list of errors that the API may return.</span></span>
 
|<span data-ttu-id="e729f-197">error.code</span><span class="sxs-lookup"><span data-stu-id="e729f-197">error.code</span></span> | <span data-ttu-id="e729f-198">error.message</span><span class="sxs-lookup"><span data-stu-id="e729f-198">error.message</span></span> |
|:----------|:--------------|
|<span data-ttu-id="e729f-199">InvalidArgument</span><span class="sxs-lookup"><span data-stu-id="e729f-199">InvalidArgument</span></span> |<span data-ttu-id="e729f-200">O argumento é inválido, ausente ou tem um formato incorreto.</span><span class="sxs-lookup"><span data-stu-id="e729f-200">The argument is invalid or missing or has an incorrect format.</span></span>|
|<span data-ttu-id="e729f-201">InvalidRequest</span><span class="sxs-lookup"><span data-stu-id="e729f-201">InvalidRequest</span></span>  |<span data-ttu-id="e729f-202">Não é possível processar a solicitação.</span><span class="sxs-lookup"><span data-stu-id="e729f-202">Cannot process the request.</span></span>|
|<span data-ttu-id="e729f-203">InvalidReference</span><span class="sxs-lookup"><span data-stu-id="e729f-203">InvalidReference</span></span>|<span data-ttu-id="e729f-204">Essa referência não é válida para a operação atual.</span><span class="sxs-lookup"><span data-stu-id="e729f-204">This reference is not valid for the current operation.</span></span>|
|<span data-ttu-id="e729f-205">InvalidBinding</span><span class="sxs-lookup"><span data-stu-id="e729f-205">InvalidBinding</span></span>  |<span data-ttu-id="e729f-206">Essa associação de objetos não é mais válida devido a atualizações anteriores.</span><span class="sxs-lookup"><span data-stu-id="e729f-206">This object binding is no longer valid due to previous updates.</span></span>|
|<span data-ttu-id="e729f-207">InvalidSelection</span><span class="sxs-lookup"><span data-stu-id="e729f-207">InvalidSelection</span></span>|<span data-ttu-id="e729f-208">A seleção atual é inválida para esta operação.</span><span class="sxs-lookup"><span data-stu-id="e729f-208">The current selection is invalid for this operation.</span></span>|
|<span data-ttu-id="e729f-209">Unauthenticated</span><span class="sxs-lookup"><span data-stu-id="e729f-209">Unauthenticated</span></span> |<span data-ttu-id="e729f-210">Informações de autenticação necessárias estão ausentes ou inválidas.</span><span class="sxs-lookup"><span data-stu-id="e729f-210">Required authentication information is either missing or invalid.</span></span>|
|<span data-ttu-id="e729f-211">AccessDenied</span><span class="sxs-lookup"><span data-stu-id="e729f-211">AccessDenied</span></span> |<span data-ttu-id="e729f-212">Você não pode realizar a operação solicitada.</span><span class="sxs-lookup"><span data-stu-id="e729f-212">You cannot perform the requested operation.</span></span>|
|<span data-ttu-id="e729f-213">ItemNotFound</span><span class="sxs-lookup"><span data-stu-id="e729f-213">ItemNotFound</span></span> |<span data-ttu-id="e729f-214">O recurso solicitado não existe.</span><span class="sxs-lookup"><span data-stu-id="e729f-214">The requested resource doesn't exist.</span></span>|
|<span data-ttu-id="e729f-215">ActivityLimitReached</span><span class="sxs-lookup"><span data-stu-id="e729f-215">ActivityLimitReached</span></span>|<span data-ttu-id="e729f-216">O limite de atividades foi alcançado.</span><span class="sxs-lookup"><span data-stu-id="e729f-216">Activity limit has been reached.</span></span>|
|<span data-ttu-id="e729f-217">GeneralException</span><span class="sxs-lookup"><span data-stu-id="e729f-217">GeneralException</span></span>|<span data-ttu-id="e729f-218">Ocorreu um erro interno ao processar a solicitação.</span><span class="sxs-lookup"><span data-stu-id="e729f-218">There was an internal error while processing the request.</span></span>|
|<span data-ttu-id="e729f-219">NotImplemented</span><span class="sxs-lookup"><span data-stu-id="e729f-219">NotImplemented</span></span>  |<span data-ttu-id="e729f-220">O recurso solicitado não foi implementado.</span><span class="sxs-lookup"><span data-stu-id="e729f-220">The requested feature isn't implemented.</span></span>|
|<span data-ttu-id="e729f-221">ServiceNotAvailable</span><span class="sxs-lookup"><span data-stu-id="e729f-221">ServiceNotAvailable</span></span>|<span data-ttu-id="e729f-222">O serviço não está disponível.</span><span class="sxs-lookup"><span data-stu-id="e729f-222">The service is unavailable.</span></span>|
|<span data-ttu-id="e729f-223">Conflict</span><span class="sxs-lookup"><span data-stu-id="e729f-223">Conflict</span></span>              |<span data-ttu-id="e729f-224">A solicitação não pôde ser processada devido a um conflito.</span><span class="sxs-lookup"><span data-stu-id="e729f-224">Request could not be processed because of a conflict.</span></span>|
|<span data-ttu-id="e729f-225">ItemAlreadyExists</span><span class="sxs-lookup"><span data-stu-id="e729f-225">ItemAlreadyExists</span></span>|<span data-ttu-id="e729f-226">O recurso que está sendo criado já existe.</span><span class="sxs-lookup"><span data-stu-id="e729f-226">The resource being created already exists.</span></span>|
|<span data-ttu-id="e729f-227">UnsupportedOperation</span><span class="sxs-lookup"><span data-stu-id="e729f-227">UnsupportedOperation</span></span>|<span data-ttu-id="e729f-228">Não há suporte para a operação.</span><span class="sxs-lookup"><span data-stu-id="e729f-228">The operation being attempted is not supported.</span></span>|
|<span data-ttu-id="e729f-229">RequestAborted</span><span class="sxs-lookup"><span data-stu-id="e729f-229">RequestAborted</span></span>|<span data-ttu-id="e729f-230">A solicitação foi anulada durante o tempo de execução.</span><span class="sxs-lookup"><span data-stu-id="e729f-230">The request was aborted during run time.</span></span>|
|<span data-ttu-id="e729f-231">ApiNotAvailable</span><span class="sxs-lookup"><span data-stu-id="e729f-231">ApiNotAvailable</span></span>|<span data-ttu-id="e729f-232">A API solicitada não está disponível.</span><span class="sxs-lookup"><span data-stu-id="e729f-232">The requested API is not available.</span></span>|
|<span data-ttu-id="e729f-233">InsertDeleteConflict</span><span class="sxs-lookup"><span data-stu-id="e729f-233">InsertDeleteConflict</span></span>|<span data-ttu-id="e729f-234">A operação de exclusão ou inserção resultou em um conflito.</span><span class="sxs-lookup"><span data-stu-id="e729f-234">The insert or delete operation attempted resulted in a conflict.</span></span>|
|<span data-ttu-id="e729f-235">InvalidOperation</span><span class="sxs-lookup"><span data-stu-id="e729f-235">InvalidOperation</span></span>|<span data-ttu-id="e729f-236">A operação é inválida no objeto.</span><span class="sxs-lookup"><span data-stu-id="e729f-236">The operation attempted is invalid on the object.</span></span>|
 
## <a name="see-also"></a><span data-ttu-id="e729f-237">Veja também</span><span class="sxs-lookup"><span data-stu-id="e729f-237">See also</span></span>
 
* [<span data-ttu-id="e729f-238">Introdução aos suplementos do Excel</span><span class="sxs-lookup"><span data-stu-id="e729f-238">Get started with Excel add-ins</span></span>](excel-add-ins-get-started-overview.md)
* [<span data-ttu-id="e729f-239">Exemplos de códigos de suplementos do Excel</span><span class="sxs-lookup"><span data-stu-id="e729f-239">Excel add-ins code samples</span></span>](https://developer.microsoft.com/office/gallery/?filterBy=Samples)
* [<span data-ttu-id="e729f-240">Conceitos de programação avançados com a API JavaScript do Excel</span><span class="sxs-lookup"><span data-stu-id="e729f-240">Advanced programming concepts with the Excel JavaScript API</span></span>](excel-add-ins-advanced-concepts.md)
* [<span data-ttu-id="e729f-241">Otimização de desempenho da API JavaScript do Excel</span><span class="sxs-lookup"><span data-stu-id="e729f-241">Excel JavaScript API performance optimization</span></span>](https://docs.microsoft.com/office/dev/add-ins/excel/performance)
* [<span data-ttu-id="e729f-242">Referência da API JavaScript do Excel</span><span class="sxs-lookup"><span data-stu-id="e729f-242">Excel JavaScript API reference</span></span>](https://docs.microsoft.com/office/dev/add-ins/reference/overview/excel-add-ins-reference-overview?view=office-js)
