---
ms.date: 05/07/2019
description: Solicite, transmita e cancele o fluxo de dados externos para sua pasta de trabalho com funções personalizadas no Excel
title: Receber e tratar dados com funções personalizadas
localization_priority: Priority
ms.openlocfilehash: 61f4d0fdaea4277faedddbe075a587fb23842c08
ms.sourcegitcommit: 5b9c2b39dfe76cabd98bf28d5287d9718788e520
ms.translationtype: HT
ms.contentlocale: pt-BR
ms.lasthandoff: 05/07/2019
ms.locfileid: "33659632"
---
# <a name="receive-and-handle-data-with-custom-functions"></a><span data-ttu-id="0e685-103">Receber e tratar dados com funções personalizadas</span><span class="sxs-lookup"><span data-stu-id="0e685-103">Receive and handle data with custom functions</span></span>

<span data-ttu-id="0e685-104">Uma das maneiras pelas quais as funções personalizadas aprimoram o poder do Excel é através do recebimento de dados de outros locais diferente da pasta de trabalho, como a Web ou um servidor (por meio de WebSockets).</span><span class="sxs-lookup"><span data-stu-id="0e685-104">One of the ways that custom functions enhance Excel's power is by receiving data from locations other than the workbook, such as the web or a server (through WebSockets).</span></span> <span data-ttu-id="0e685-105">As funções personalizadas podem solicitar dados por meio de XHR e solicitações `fetch`, bem como transmitir esses dados em tempo real.</span><span class="sxs-lookup"><span data-stu-id="0e685-105">Custom functions can request data through XHR and fetch requests as well as stream this data in real time.</span></span>

[!include[Excel custom functions note](../includes/excel-custom-functions-note.md)]

<span data-ttu-id="0e685-106">A documentação a seguir ilustra alguns exemplos de solicitações da web, mas para criar uma função de transmissão para você, experimente o [Tutorial de funções personalizadas](https://docs.microsoft.com/office/dev/add-ins/tutorials/excel-tutorial-create-custom-functions?tabs=excel-windows).</span><span class="sxs-lookup"><span data-stu-id="0e685-106">The documentation below illustrates some samples of web requests, but to build a streaming function for yourself, try the [Custom functions tutorial](https://docs.microsoft.com/office/dev/add-ins/tutorials/excel-tutorial-create-custom-functions?tabs=excel-windows).</span></span>

## <a name="functions-that-return-data-from-external-sources"></a><span data-ttu-id="0e685-107">Funções que retornam os dados de fontes externas</span><span class="sxs-lookup"><span data-stu-id="0e685-107">Functions that return data from external sources</span></span>

<span data-ttu-id="0e685-108">Se uma função personalizada recupera dados de uma fonte externa como na web, ela deve:</span><span class="sxs-lookup"><span data-stu-id="0e685-108">If a custom function retrieves data from an external source such as the web, it must:</span></span>

1. <span data-ttu-id="0e685-109">Retornar uma Promise do JavaScript para o Excel.</span><span class="sxs-lookup"><span data-stu-id="0e685-109">Return a JavaScript Promise to Excel.</span></span>
2. <span data-ttu-id="0e685-110">Resolva a promessa com o uso da função retorno de chamada de valor final.</span><span class="sxs-lookup"><span data-stu-id="0e685-110">Resolve the Promise with the final value using the callback function.</span></span>

<span data-ttu-id="0e685-111">É possível solicitar dados externos através de uma API como [ `Fetch` ](https://developer.mozilla.org/pt-BR/docs/Web/API/Fetch_API) ou usando `XmlHttpRequest` [(XHR)](https://developer.mozilla.org/pt-BR/docs/Web/API/XMLHttpRequest), uma API Web padrão que envia solicitações HTTP para interagir com os servidores.</span><span class="sxs-lookup"><span data-stu-id="0e685-111">You can request external data through an API like [`Fetch`](https://developer.mozilla.org/pt-BR/docs/Web/API/Fetch_API) or by using `XmlHttpRequest` [(XHR)](https://developer.mozilla.org/pt-BR/docs/Web/API/XMLHttpRequest), a standard web API that issues HTTP requests to interact with servers.</span></span>

<span data-ttu-id="0e685-112">No tempo de execução das funções personalizadas, o XHR implementa medidas de segurança adicionais solicitando uma [Política de mesma origem](https://developer.mozilla.org/pt-BR/docs/Web/Security/Same-origin_policy) ou um simples [CORS](https://www.w3.org/TR/cors/).</span><span class="sxs-lookup"><span data-stu-id="0e685-112">Within custom functions runtime, XHR implements additional security measures by requiring [Same Origin Policy](https://developer.mozilla.org/pt-BR/docs/Web/Security/Same-origin_policy) and simple [CORS](https://www.w3.org/TR/cors/).</span></span>

<span data-ttu-id="0e685-113">Observe que uma implementação CORS simples não pode usar cookies e é compatível apenas com métodos simples (GET, HEAD, POST).</span><span class="sxs-lookup"><span data-stu-id="0e685-113">Note that a simple CORS implementation cannot use cookies and only supports simple methods (GET, HEAD, POST).</span></span> <span data-ttu-id="0e685-114">A CORS simples aceita cabeçalhos simples com nomes de campos `Accept`, `Accept-Language`, `Content-Language`.</span><span class="sxs-lookup"><span data-stu-id="0e685-114">Simple CORS accepts simple headers with field names `Accept`, `Accept-Language`, `Content-Language`.</span></span> <span data-ttu-id="0e685-115">Você também pode usar um cabeçalho de tipo de conteúdo no CORS simples, desde que o tipo de conteúdo seja `application/x-www-form-urlencoded`, `text/plain`, ou `multipart/form-data`.</span><span class="sxs-lookup"><span data-stu-id="0e685-115">You can also use a Content-Type header in simple CORS, provided that the content type is `application/x-www-form-urlencoded`, `text/plain`, or `multipart/form-data`.</span></span>

### <a name="xhr-example"></a><span data-ttu-id="0e685-116">Exemplo de XHR</span><span class="sxs-lookup"><span data-stu-id="0e685-116">XHR example</span></span>

<span data-ttu-id="0e685-117">No código de exemplo a seguir, a função **getTemperature** chama a função sendWebRequest  para obter a temperatura de uma área específica, de acordo com a ID do termômetro.</span><span class="sxs-lookup"><span data-stu-id="0e685-117">In the following code sample, the **getTemperature** function calls the sendWebRequest function to get the temperature of a particular area based on thermometer ID.</span></span> <span data-ttu-id="0e685-118">A função sendWebRequest usa XHR para emitir uma solicitação GET para um ponto de extremidade que fornece os dados.</span><span class="sxs-lookup"><span data-stu-id="0e685-118">The sendWebRequest function uses XHR to issue a GET request to an endpoint that can provide the data.</span></span>

```js
/**
 * Receives a temperature from an online source.
 * @customfunction
 * @param {number} thermometerID Identification number of the thermometer.
 */
function getTemperature(thermometerID) {
  return new Promise(function(setResult) {
      sendWebRequest(thermometerID, function(data){
          storeLastTemperature(thermometerID, data.temperature);
          setResult(data.temperature);
      });
  });
}

// Helper method that uses Office's implementation of XMLHttpRequest in the JavaScript runtime for custom functions.  
function sendWebRequest(thermometerID, data) {
    var xhttp = new XMLHttpRequest();
    xhttp.onreadystatechange = function() {
        if (this.readyState == 4 && this.status == 200) {
           data.temperature = JSON.parse(xhttp.responseText).temperature
        };

        //set Content-Type to application/text. Application/json is not currently supported with Simple CORS
        xhttp.setRequestHeader("Content-Type", "application/text");
        xhttp.open("GET", "https://contoso.com/temperature/" + thermometerID), true)
        xhttp.send();  
    }
}

CustomFunctions.associate("GETTEMPERATURE", getTemperature);
```

<span data-ttu-id="0e685-119">Para outro exemplo de solicitação XHR com mais contexto, confira a função`getFile` dentro [deste arquivo](https://github.com/OfficeDev/Office-Add-in-JavaScript-FileDownload/blob/master/FileDownloadSampleWeb/Home.js) no repositório Github [Office-Add-in-JavaScript-FileDownload](https://github.com/OfficeDev/Office-Add-in-JavaScript-FileDownload).</span><span class="sxs-lookup"><span data-stu-id="0e685-119">For another sample of an XHR request with more context, see the `getFile` function within [this file](https://github.com/OfficeDev/Office-Add-in-JavaScript-FileDownload/blob/master/FileDownloadSampleWeb/Home.js) in the [Office-Add-in-JavaScript-FileDownload](https://github.com/OfficeDev/Office-Add-in-JavaScript-FileDownload) Github repository.</span></span>

### <a name="fetch-example"></a><span data-ttu-id="0e685-120">Exemplo de busca</span><span class="sxs-lookup"><span data-stu-id="0e685-120">Fetch example</span></span>

<span data-ttu-id="0e685-121">No seguinte exemplo de código, a função `stockPriceStream` usa um símbolo de cotação da bolsa para acessar o preço de uma ação a cada 1000 milissegundos.</span><span class="sxs-lookup"><span data-stu-id="0e685-121">In the following code sample, the stockPriceStream function uses a stock ticker symbol to get the price of a stock every 1000 milliseconds.</span></span> <span data-ttu-id="0e685-122">Para saber mais sobre este exemplo, confira o [Tutorial de funções personalizadas](https://docs.microsoft.com/office/dev/add-ins/tutorials/excel-tutorial-create-custom-functions?tabs=excel-windows#create-a-streaming-asynchronous-custom-function).</span><span class="sxs-lookup"><span data-stu-id="0e685-122">For more details about this sample and to get the accompanying JSON, see the [Custom functions tutorial](https://docs.microsoft.com/office/dev/add-ins/tutorials/excel-tutorial-create-custom-functions?tabs=excel-windows#create-a-streaming-asynchronous-custom-function).</span></span>

```js
/**
 * Streams a stock price.
 * @customfunction 
 * @param {string} ticker Stock ticker.
 * @param {CustomFunctions.StreamingInvocation<number>} invocation Invocation parameter necessary for streaming functions.
 */
function stockPriceStream(ticker, invocation) {
    var updateFrequency = 1000 /* milliseconds*/;
    var isPending = false;

    var timer = setInterval(function() {
        // If there is already a pending request, skip this iteration:
        if (isPending) {
            return;
        }

        var url = "https://api.iextrading.com/1.0/stock/" + ticker + "/price";
        isPending = true;

        fetch(url)
            .then(function(response) {
                return response.text();
            })
            .then(function(text) {
                invocation.setResult(parseFloat(text));
            })
            .catch(function(error) {
                invocation.setResult(error);
            })
            .then(function() {
                isPending = false;
            });
    }, updateFrequency);

    invocation.onCanceled = () => {
        clearInterval(timer);
    };
}

CustomFunctions.associate("STOCKPRICESTREAM", stockPriceStream);
```

## <a name="receive-data-via-websockets"></a><span data-ttu-id="0e685-123">Receber dados por meio de WebSockets</span><span class="sxs-lookup"><span data-stu-id="0e685-123">Receiving data via WebSockets</span></span>

<span data-ttu-id="0e685-124">Em uma função personalizada, é possível usar WebSockets para trocar dados por meio de uma conexão persistente com um servidor.</span><span class="sxs-lookup"><span data-stu-id="0e685-124">Within a custom function, you can use WebSockets to exchange data over a persistent connection with a server.</span></span> <span data-ttu-id="0e685-125">Usando WebSockets, a função personalizada pode abrir uma conexão com um servidor e, em seguida, receber mensagens do servidor automaticamente, quando determinados eventos ocorrerem, sem precisar consultar explicitamente os dados do servidor.</span><span class="sxs-lookup"><span data-stu-id="0e685-125">By using WebSockets, your custom function can open a connection with a server and then automatically receive messages from the server when certain events occur, without having to explicitly poll the server for data.</span></span>

### <a name="websockets-example"></a><span data-ttu-id="0e685-126">Exemplo de WebSockets</span><span class="sxs-lookup"><span data-stu-id="0e685-126">WebSockets example</span></span>

<span data-ttu-id="0e685-127">O código de exemplo a seguir estabelece uma conexão WebSocket e registra cada mensagem de entrada do servidor.</span><span class="sxs-lookup"><span data-stu-id="0e685-127">The following code sample establishes a WebSocket connection and then logs each incoming message from the server.</span></span>

```js
let ws = new WebSocket('wss://bundles.office.com');

ws.onmessage(message) {
    console.log(`Received: ${message}`);
}

ws.onerror(error){
    console.err(`Failed: ${error}`);
}
```

## <a name="stream-and-cancel-functions"></a><span data-ttu-id="0e685-128">Funções stream e cancel</span><span class="sxs-lookup"><span data-stu-id="0e685-128">Stream and cancel functions</span></span>

<span data-ttu-id="0e685-129">Funções personalizadas de streaming permitem a saída de dados para células que atualizam repetidamente, sem a necessidade de um usuário explicitamente atualizar coisa alguma.</span><span class="sxs-lookup"><span data-stu-id="0e685-129">Streaming custom functions enable you to output data to cells repeatedly over time, without requiring a user to explicitly request data refresh.</span></span>

<span data-ttu-id="0e685-130">Funções personalizadas cancelable permitem com que você cancele a execução de uma função personalizada de streaming para reduzir o consumo de banda larga, memória de trabalho e carregamento de CPU.</span><span class="sxs-lookup"><span data-stu-id="0e685-130">In some situations, you may need to cancel the execution of a streaming custom function to reduce its bandwidth consumption, working memory, and CPU load.</span></span>

<span data-ttu-id="0e685-131">Para declarar uma função como streaming ou cancelable, use as marcas de comentário JSDOC `@stream` ou `@cancelable`. </span><span class="sxs-lookup"><span data-stu-id="0e685-131">To declare a function as streaming or cancelable, use the JSDOC comment tags `@stream` or `@cancelable`.</span></span>

### <a name="using-an-invocation-parameter"></a><span data-ttu-id="0e685-132">Usando um parâmetro de invocação</span><span class="sxs-lookup"><span data-stu-id="0e685-132">Using an invocation parameter</span></span>

<span data-ttu-id="0e685-133">O parâmetro `invocation` é o último parâmetro de qualquer função personalizada por padrão.</span><span class="sxs-lookup"><span data-stu-id="0e685-133">The `invocation` parameter is the last parameter of any custom function by default.</span></span> <span data-ttu-id="0e685-134">O parâmetro `invocation` fornece um contexto sobre a célula (como o seu endereço) e também permite com que você use os métodos `setResult` e `onCanceled`.</span><span class="sxs-lookup"><span data-stu-id="0e685-134">The `invocation` parameter gives context about the cell (such as its address) and also allows you to use `setResult` and `onCanceled` methods.</span></span> <span data-ttu-id="0e685-135">Esses métodos definem o que uma função faz quando a função transmite (`setResult`) ou é cancelada (`onCanceled`).</span><span class="sxs-lookup"><span data-stu-id="0e685-135">These methods define what a function does when the function streams (`setResult`) or is canceled (`onCanceled`).</span></span>

<span data-ttu-id="0e685-136">Se você estiver usando o TypeScript, o manipulador de invocações deve ser do tipo `CustomFunctions.StreamingInvocation` ou `CustomFunctions.CancelableInvocation`.</span><span class="sxs-lookup"><span data-stu-id="0e685-136">If you're using TypeScript, the invocation handler needs to be of type `CustomFunctions.StreamingInvocation` or `CustomFunctions.CancelableInvocation`.</span></span>

### <a name="streaming-and-cancelable-function-example"></a><span data-ttu-id="0e685-137">Exemplo das funções streaming e cancelable</span><span class="sxs-lookup"><span data-stu-id="0e685-137">Streaming and cancelable function example</span></span>
<span data-ttu-id="0e685-138">O exemplo a seguir é uma função personalizada que adiciona um número ao resultado a cada segundo.</span><span class="sxs-lookup"><span data-stu-id="0e685-138">The following code sample is a custom function that adds a number to the result every second.</span></span> <span data-ttu-id="0e685-139">Observe o seguinte sobre este código:</span><span class="sxs-lookup"><span data-stu-id="0e685-139">Note the following about this code:</span></span>

- <span data-ttu-id="0e685-140">O Excel exibe cada valor novo automaticamente usando o método `setResult`.</span><span class="sxs-lookup"><span data-stu-id="0e685-140">Excel displays each new value automatically using the `setResult` callback.</span></span>
- <span data-ttu-id="0e685-141">O segundo parâmetro de entrada, invocação, não é exibido para os usuários finais no Excel quando eles selecionam a função no menu de preenchimento automático.</span><span class="sxs-lookup"><span data-stu-id="0e685-141">The second input parameter, , is not displayed to end users in Excel when they select the function from the autocomplete menu.</span></span>
- <span data-ttu-id="0e685-142">O retorno de chamada `onCanceled` define a função que é executada quando a função é cancelada.</span><span class="sxs-lookup"><span data-stu-id="0e685-142">The `onCanceled` callback defines the function that executes when the function is canceled.</span></span>

```js
/**
 * Increments a value once a second.
 * @customfunction
 * @param {number} incrementBy Amount to increment.
 * @param {CustomFunctions.StreamingInvocation<number>} invocation Invocation parameter necessary for streaming functions.
 */
function increment(incrementBy, invocation) {
  let result = 0;
  const timer = setInterval(() => {
    result += incrementBy;
    invocation.setResult(result);
  }, 1000);

  invocation.onCanceled = function(){
    clearInterval(timer);
    }
}
CustomFunctions.associate("INCREMENT", increment);
```

>[!NOTE]
> <span data-ttu-id="0e685-143">O Excel cancela a execução de uma função nas seguintes situações:</span><span class="sxs-lookup"><span data-stu-id="0e685-143">Excel cancels the execution of a function in the following situations:</span></span>
>
> - <span data-ttu-id="0e685-144">Quando o usuário edita ou exclui uma célula que faz referência à função.</span><span class="sxs-lookup"><span data-stu-id="0e685-144">When the user edits or deletes a cell that references the function.</span></span>
> - <span data-ttu-id="0e685-145">Quando é alterado um dos argumentos (entradas) para a função.</span><span class="sxs-lookup"><span data-stu-id="0e685-145">When one of the arguments (inputs) for the function changes.</span></span> <span data-ttu-id="0e685-146">Nesse caso, uma nova chamada de função é disparada, seguindo o cancelamento.</span><span class="sxs-lookup"><span data-stu-id="0e685-146">In this case, a new function call is triggered following the cancellation.</span></span>
> - <span data-ttu-id="0e685-147">Quando o usuário aciona manualmente um recálculo.</span><span class="sxs-lookup"><span data-stu-id="0e685-147">When the user triggers recalculation manually.</span></span> <span data-ttu-id="0e685-148">Nesse caso, uma nova chamada de função é disparada, seguindo o cancelamento.</span><span class="sxs-lookup"><span data-stu-id="0e685-148">In this case, a new function call is triggered following the cancellation.</span></span>

## <a name="next-steps"></a><span data-ttu-id="0e685-149">Próximas etapas</span><span class="sxs-lookup"><span data-stu-id="0e685-149">Next steps</span></span>

* <span data-ttu-id="0e685-150">Saiba mais sobre [diferentes tipos de parâmetros que as suas funções podem usar](custom-functions-parameter-options.md).</span><span class="sxs-lookup"><span data-stu-id="0e685-150">Learn about [different parameter types your functions can use](custom-functions-parameter-options.md).</span></span>
* <span data-ttu-id="0e685-151">Descubra como [agrupar várias chamadas de API](custom-functions-batching.md).</span><span class="sxs-lookup"><span data-stu-id="0e685-151">Discover how to [batch multiple API calls](custom-functions-batching.md).</span></span>

## <a name="see-also"></a><span data-ttu-id="0e685-152">Confira também</span><span class="sxs-lookup"><span data-stu-id="0e685-152">See also</span></span>

* [<span data-ttu-id="0e685-153">Valores voláteis nas funções</span><span class="sxs-lookup"><span data-stu-id="0e685-153">Volatile values in functions</span></span>](custom-functions-volatile.md)
* [<span data-ttu-id="0e685-154">Criar metadados JSON para funções personalizadas</span><span class="sxs-lookup"><span data-stu-id="0e685-154">Create JSON metadata for custom functions</span></span>](custom-functions-json-autogeneration.md)
* [<span data-ttu-id="0e685-155">Metadados de funções personalizadas</span><span class="sxs-lookup"><span data-stu-id="0e685-155">Custom functions metadata</span></span>](custom-functions-json.md)
* [<span data-ttu-id="0e685-156">Tempo de execução de funções personalizadas do Excel</span><span class="sxs-lookup"><span data-stu-id="0e685-156">Runtime for Excel custom functions</span></span>](custom-functions-runtime.md)
* <span data-ttu-id="0e685-157">[Práticas recomendadas de funções personalizadas](custom-functions-best-practices.md).</span><span class="sxs-lookup"><span data-stu-id="0e685-157">[Custom functions best practices](custom-functions-best-practices.md)</span></span>
* [<span data-ttu-id="0e685-158">Criar funções personalizadas no Excel</span><span class="sxs-lookup"><span data-stu-id="0e685-158">Create custom functions in Excel</span></span>](custom-functions-overview.md)
* [<span data-ttu-id="0e685-159">Tutorial de funções personalizadas do Excel</span><span class="sxs-lookup"><span data-stu-id="0e685-159">Excel custom functions tutorial</span></span>](../tutorials/excel-tutorial-create-custom-functions.md)
