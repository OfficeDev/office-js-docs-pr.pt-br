---
ms.date: 06/21/2019
description: Solicite, transmita e cancele o fluxo de dados externos para sua pasta de trabalho com funções personalizadas no Excel
title: Receber e tratar dados com funções personalizadas
localization_priority: Priority
ms.openlocfilehash: 39be2f0913e2eee4b1e5e7d5f704a47dee279cf5
ms.sourcegitcommit: 382e2735a1295da914f2bfc38883e518070cec61
ms.translationtype: HT
ms.contentlocale: pt-BR
ms.lasthandoff: 06/21/2019
ms.locfileid: "35128252"
---
# <a name="receive-and-handle-data-with-custom-functions"></a><span data-ttu-id="da76c-103">Receber e tratar dados com funções personalizadas</span><span class="sxs-lookup"><span data-stu-id="da76c-103">Receive and handle data with custom functions</span></span>

<span data-ttu-id="da76c-104">Uma das maneiras pelas quais as funções personalizadas aprimoram o poder do Excel é através do recebimento de dados de outros locais diferente da pasta de trabalho, como a Web ou um servidor (por meio de WebSockets).</span><span class="sxs-lookup"><span data-stu-id="da76c-104">One of the ways that custom functions enhances Excel's power is by receiving data from locations other than the workbook, such as the web or a server (through WebSockets).</span></span> <span data-ttu-id="da76c-105">As funções personalizadas podem solicitar dados por meio de XHR e solicitações `fetch`, bem como transmitir esses dados em tempo real.</span><span class="sxs-lookup"><span data-stu-id="da76c-105">Custom functions can request data through XHR and `fetch` requests as well as stream this data in real time.</span></span>

[!include[Excel custom functions note](../includes/excel-custom-functions-note.md)]

<span data-ttu-id="da76c-106">A documentação a seguir ilustra alguns exemplos de solicitações da web, mas para criar uma função de transmissão para você, experimente o [Tutorial de funções personalizadas](../tutorials/excel-tutorial-create-custom-functions.md).</span><span class="sxs-lookup"><span data-stu-id="da76c-106">The documentation below illustrates some samples of web requests, but to build a streaming function for yourself, try the [Custom functions tutorial](../tutorials/excel-tutorial-create-custom-functions.md).</span></span>

## <a name="functions-that-return-data-from-external-sources"></a><span data-ttu-id="da76c-107">Funções que retornam os dados de fontes externas</span><span class="sxs-lookup"><span data-stu-id="da76c-107">Functions that return data from external sources</span></span>

<span data-ttu-id="da76c-108">Se uma função personalizada recupera dados de uma fonte externa como na web, ela deve:</span><span class="sxs-lookup"><span data-stu-id="da76c-108">If a custom function retrieves data from an external source such as the web, it must:</span></span>

1. <span data-ttu-id="da76c-109">Retornar uma Promise do JavaScript para o Excel.</span><span class="sxs-lookup"><span data-stu-id="da76c-109">Return a JavaScript Promise to Excel.</span></span>
2. <span data-ttu-id="da76c-110">Resolva a promessa com o uso da função retorno de chamada de valor final.</span><span class="sxs-lookup"><span data-stu-id="da76c-110">Resolve the Promise with the final value using the callback function.</span></span>

<span data-ttu-id="da76c-111">É possível solicitar dados externos através de uma API como [ `Fetch` ](https://developer.mozilla.org/en-US/docs/Web/API/Fetch_API) ou usando `XmlHttpRequest` [(XHR)](https://developer.mozilla.org/en-US/docs/Web/API/XMLHttpRequest), uma API Web padrão que envia solicitações HTTP para interagir com os servidores.</span><span class="sxs-lookup"><span data-stu-id="da76c-111">You can request external data through an API like [`Fetch`](https://developer.mozilla.org/en-US/docs/Web/API/Fetch_API) or by using `XmlHttpRequest` [(XHR)](https://developer.mozilla.org/en-US/docs/Web/API/XMLHttpRequest), a standard web API that issues HTTP requests to interact with servers.</span></span>

<span data-ttu-id="da76c-112">No tempo de execução das funções personalizadas, o XHR implementa medidas de segurança adicionais solicitando uma [Política de mesma origem](https://developer.mozilla.org/en-US/docs/Web/Security/Same-origin_policy) ou um simples [CORS](https://www.w3.org/TR/cors/).</span><span class="sxs-lookup"><span data-stu-id="da76c-112">Within custom functions runtime, XHR implements additional security measures by requiring [Same Origin Policy](https://developer.mozilla.org/en-US/docs/Web/Security/Same-origin_policy) and simple [CORS](https://www.w3.org/TR/cors/).</span></span>

<span data-ttu-id="da76c-113">Observe que uma implementação CORS simples não pode usar cookies e é compatível apenas com métodos simples (GET, HEAD, POST).</span><span class="sxs-lookup"><span data-stu-id="da76c-113">Note that a simple CORS implementation cannot use cookies and only supports simple methods (GET, HEAD, POST).</span></span> <span data-ttu-id="da76c-114">A CORS simples aceita cabeçalhos simples com nomes de campos `Accept`, `Accept-Language`, `Content-Language`.</span><span class="sxs-lookup"><span data-stu-id="da76c-114">Simple CORS accepts simple headers with field names `Accept`, `Accept-Language`, `Content-Language`.</span></span> <span data-ttu-id="da76c-115">Você também pode usar um cabeçalho de tipo de conteúdo no CORS simples, desde que o tipo de conteúdo seja `application/x-www-form-urlencoded`, `text/plain`, ou `multipart/form-data`.</span><span class="sxs-lookup"><span data-stu-id="da76c-115">You can also use a Content-Type header in simple CORS, provided that the content type is `application/x-www-form-urlencoded`, `text/plain`, or `multipart/form-data`.</span></span>

### <a name="xhr-example"></a><span data-ttu-id="da76c-116">Exemplo de XHR</span><span class="sxs-lookup"><span data-stu-id="da76c-116">XHR example</span></span>

<span data-ttu-id="da76c-117">No código de exemplo a seguir, a função **getTemperature** chama a função sendWebRequest  para obter a temperatura de uma área específica, de acordo com a ID do termômetro.</span><span class="sxs-lookup"><span data-stu-id="da76c-117">In the following code sample, the **getTemperature** function calls the sendWebRequest function to get the temperature of a particular area based on thermometer ID.</span></span> <span data-ttu-id="da76c-118">A função sendWebRequest usa XHR para emitir uma solicitação GET para um ponto de extremidade que fornece os dados.</span><span class="sxs-lookup"><span data-stu-id="da76c-118">The sendWebRequest function uses XHR to issue a GET request to an endpoint that can provide the data.</span></span>

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

<span data-ttu-id="da76c-119">Para outro exemplo de solicitação XHR com mais contexto, confira a função`getFile` dentro [deste arquivo](https://github.com/OfficeDev/Office-Add-in-JavaScript-FileDownload/blob/master/FileDownloadSampleWeb/Home.js) no repositório Github [Office-Add-in-JavaScript-FileDownload](https://github.com/OfficeDev/Office-Add-in-JavaScript-FileDownload).</span><span class="sxs-lookup"><span data-stu-id="da76c-119">For another sample of an XHR request with more context, see the `getFile` function within [this file](https://github.com/OfficeDev/Office-Add-in-JavaScript-FileDownload/blob/master/FileDownloadSampleWeb/Home.js) in the [Office-Add-in-JavaScript-FileDownload](https://github.com/OfficeDev/Office-Add-in-JavaScript-FileDownload) Github repository.</span></span>

### <a name="fetch-example"></a><span data-ttu-id="da76c-120">Exemplo de busca</span><span class="sxs-lookup"><span data-stu-id="da76c-120">Fetch example</span></span>

<span data-ttu-id="da76c-121">No seguinte exemplo de código, a função `stockPriceStream` usa um símbolo de cotação da bolsa para acessar o preço de uma ação a cada 1000 milissegundos.</span><span class="sxs-lookup"><span data-stu-id="da76c-121">In the following code sample, the `stockPriceStream` function uses a stock ticker symbol to get the price of a stock every 1000 milliseconds.</span></span> <span data-ttu-id="da76c-122">Para saber mais sobre este exemplo, confira o [Tutorial de funções personalizadas](../tutorials/excel-tutorial-create-custom-functions.md#create-a-streaming-asynchronous-custom-function).</span><span class="sxs-lookup"><span data-stu-id="da76c-122">For more details about this sample, see the [Custom functions tutorial](../tutorials/excel-tutorial-create-custom-functions.md#create-a-streaming-asynchronous-custom-function).</span></span>

> [!NOTE]
> <span data-ttu-id="da76c-123">O código a seguir solicita uma cotação de ações usando a API de Negociação IEX.</span><span class="sxs-lookup"><span data-stu-id="da76c-123">The following code requests a stock quote using the IEX Trading API.</span></span> <span data-ttu-id="da76c-124">Antes de executar o código, você precisará [criar uma conta gratuita com o IEX Cloud](https://iexcloud.io/) para poder obter o token da API necessário na solicitação de API.</span><span class="sxs-lookup"><span data-stu-id="da76c-124">Before you can run the code, you'll need to [create a free account with IEX Cloud](https://iexcloud.io/) so that you can get the API token that's required in the API request.</span></span>

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

        //Note: In the following line, replace <YOUR_TOKEN_HERE> with the API token that you've obtained through your IEX Cloud account.
        var url = "https://cloud.iexapis.com/stable/stock/" + ticker + "/quote/latestPrice?token=<YOUR_TOKEN_HERE>"
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

## <a name="receive-data-via-websockets"></a><span data-ttu-id="da76c-125">Receber dados por meio de WebSockets</span><span class="sxs-lookup"><span data-stu-id="da76c-125">Receive data via WebSockets</span></span>

<span data-ttu-id="da76c-126">Em uma função personalizada, é possível usar WebSockets para trocar dados por meio de uma conexão persistente com um servidor.</span><span class="sxs-lookup"><span data-stu-id="da76c-126">Within a custom function, you can use WebSockets to exchange data over a persistent connection with a server.</span></span> <span data-ttu-id="da76c-127">Usando WebSockets, a função personalizada pode abrir uma conexão com um servidor e, em seguida, receber mensagens do servidor automaticamente, quando determinados eventos ocorrerem, sem precisar consultar explicitamente os dados do servidor.</span><span class="sxs-lookup"><span data-stu-id="da76c-127">By using WebSockets, your custom function can open a connection with a server and then automatically receive messages from the server when certain events occur, without having to explicitly poll the server for data.</span></span>

### <a name="websockets-example"></a><span data-ttu-id="da76c-128">Exemplo de WebSockets</span><span class="sxs-lookup"><span data-stu-id="da76c-128">WebSockets example</span></span>

<span data-ttu-id="da76c-129">O código de exemplo a seguir estabelece uma conexão WebSocket e registra cada mensagem de entrada do servidor.</span><span class="sxs-lookup"><span data-stu-id="da76c-129">The following code sample establishes a WebSocket connection and then logs each incoming message from the server.</span></span>

```js
let ws = new WebSocket('wss://bundles.office.com');

ws.onmessage(message) {
    console.log(`Received: ${message}`);
}

ws.onerror(error){
    console.err(`Failed: ${error}`);
}
```

## <a name="make-a-streaming-function"></a><span data-ttu-id="da76c-130">Faça uma função de streaming</span><span class="sxs-lookup"><span data-stu-id="da76c-130">Make a streaming function</span></span>

<span data-ttu-id="da76c-131">Funções personalizadas de streaming permitem a saída de dados para células que atualizam repetidamente, sem a necessidade de um usuário explicitamente atualizar coisa alguma.</span><span class="sxs-lookup"><span data-stu-id="da76c-131">Streaming custom functions enable you to output data to cells that updates repeatedly, without requiring a user to explicitly refresh anything.</span></span> <span data-ttu-id="da76c-132">Isso pode ser útil para verificar dados ativos de um serviço online, como a função no [tutorial de funções personalizadas](../tutorials/excel-tutorial-create-custom-functions.md).</span><span class="sxs-lookup"><span data-stu-id="da76c-132">This can be useful to check live data from a service online, like the function in [the custom functions tutorial](../tutorials/excel-tutorial-create-custom-functions.md).</span></span>

<span data-ttu-id="da76c-133">Para declarar uma função de streaming, use a tag `@stream` de comentário JSDoc.</span><span class="sxs-lookup"><span data-stu-id="da76c-133">To declare a streaming function, use the JSDoc comment tag `@stream`.</span></span> <span data-ttu-id="da76c-134">Para alertar os usuários para o fato de que sua função pode ser reavaliada com base em novas informações, considere colocar fluxo ou outro texto para indicar isso no nome ou na descrição de sua função.</span><span class="sxs-lookup"><span data-stu-id="da76c-134">To alert users to the fact that your function may re-evaluate based on new information, consider putting stream or other wording to indicate this in the name or description of your function.</span></span>

<span data-ttu-id="da76c-135">O exemplo a seguir mostra uma função de streaming que aumenta um determinado número a cada segundo em um valor especificado por você.</span><span class="sxs-lookup"><span data-stu-id="da76c-135">The following example shows a streaming function which increases a given number every second by an amount you specify.</span></span>

```JS
/**
 * Increments a value once a second.
 * @customfunction INC increment
 * @param {number} incrementBy Amount to increment
 * @param {CustomFunctions.StreamingInvocation<number>} invocation
 */
function increment(incrementBy, invocation) {
  let result = 0;
  const timer = setInterval(() => {
    result += incrementBy;
    invocation.setResult(result);
  }, 1000);

  invocation.onCanceled = () => {
    clearInterval(timer);
  };
}
CustomFunctions.associate("INC", increment);
```

>[!NOTE]
> <span data-ttu-id="da76c-136">Observe que há também uma categoria de funções chamada de funções canceláveis, que *não* estão relacionadas a funções de streaming.</span><span class="sxs-lookup"><span data-stu-id="da76c-136">Note that there are also a category of functions called cancelable functions, which are *not* related to streaming functions.</span></span> <span data-ttu-id="da76c-137">Versões anteriores de funções personalizadas exigiam que você declarasse `"cancelable": true` e `"streaming": true` em JSON escrito à mão.</span><span class="sxs-lookup"><span data-stu-id="da76c-137">Previous versions of custom functions required you to declare `"cancelable": true` and `"streaming": true` in JSON written by hand.</span></span> <span data-ttu-id="da76c-138">Desde a introdução de metadados autogerados, somente as funções personalizadas assíncronas que retornam um único valor são canceláveis.</span><span class="sxs-lookup"><span data-stu-id="da76c-138">Since the introduction of autogenerated metadata, only asynchronous custom functions which return one value are cancelable.</span></span> <span data-ttu-id="da76c-139">Funções canceláveis permitem que uma solicitação da Web seja encerrada no meio de uma solicitação, usando um [`CancelableInvocation`](/javascript/api/custom-functions-runtime/customfunctions.cancelableinvocation) para decidir o que fazer após o cancelamento.</span><span class="sxs-lookup"><span data-stu-id="da76c-139">Cancelable functions allow a web request to be terminated in the middle of a request, using a [`CancelableInvocation`](/javascript/api/custom-functions-runtime/customfunctions.cancelableinvocation) to decide what to do upon cancellation.</span></span> <span data-ttu-id="da76c-140">Declare uma função cancelável usando a tag `@cancelable`.</span><span class="sxs-lookup"><span data-stu-id="da76c-140">Declare a cancelable function using the tag `@cancelable`.</span></span>

### <a name="using-an-invocation-parameter"></a><span data-ttu-id="da76c-141">Usando um parâmetro de invocação</span><span class="sxs-lookup"><span data-stu-id="da76c-141">Using an invocation parameter</span></span>

<span data-ttu-id="da76c-142">O parâmetro `invocation` é o último parâmetro de qualquer função personalizada por padrão.</span><span class="sxs-lookup"><span data-stu-id="da76c-142">The `invocation` parameter is the last parameter of any custom function by default.</span></span> <span data-ttu-id="da76c-143">O parâmetro `invocation` fornece um contexto sobre a célula (como o seu endereço) e também permite com que você use os métodos `setResult` e `onCanceled`.</span><span class="sxs-lookup"><span data-stu-id="da76c-143">The `invocation` parameter gives context about the cell (such as its address) and also allows you to use `setResult` and `onCanceled` methods.</span></span> <span data-ttu-id="da76c-144">Esses métodos definem o que uma função faz quando a função transmite (`setResult`) ou é cancelada (`onCanceled`).</span><span class="sxs-lookup"><span data-stu-id="da76c-144">These methods define what a function does when the function streams (`setResult`) or is canceled (`onCanceled`).</span></span>

<span data-ttu-id="da76c-145">Se você estiver usando o TypeScript, o manipulador de invocações deve ser do tipo `CustomFunctions.StreamingInvocation` ou `CustomFunctions.CancelableInvocation`.</span><span class="sxs-lookup"><span data-stu-id="da76c-145">If you're using TypeScript, the invocation handler needs to be of type `CustomFunctions.StreamingInvocation` or `CustomFunctions.CancelableInvocation`.</span></span>

### <a name="streaming-and-cancelable-function-example"></a><span data-ttu-id="da76c-146">Exemplo das funções streaming e cancelable</span><span class="sxs-lookup"><span data-stu-id="da76c-146">Streaming and cancelable function example</span></span>
<span data-ttu-id="da76c-147">O exemplo a seguir é uma função personalizada que adiciona um número ao resultado a cada segundo.</span><span class="sxs-lookup"><span data-stu-id="da76c-147">The following code sample is a custom function that adds a number to the result every second.</span></span> <span data-ttu-id="da76c-148">Observe o seguinte sobre este código:</span><span class="sxs-lookup"><span data-stu-id="da76c-148">Note the following about this code:</span></span>

- <span data-ttu-id="da76c-149">O Excel exibe cada valor novo automaticamente usando o método `setResult`.</span><span class="sxs-lookup"><span data-stu-id="da76c-149">Excel displays each new value automatically using the `setResult` method.</span></span>
- <span data-ttu-id="da76c-150">O segundo parâmetro de entrada, invocação, não é exibido para os usuários finais no Excel quando eles selecionam a função no menu de preenchimento automático.</span><span class="sxs-lookup"><span data-stu-id="da76c-150">The second input parameter, invocation, is not displayed to end users in Excel when they select the function from the autocomplete menu.</span></span>
- <span data-ttu-id="da76c-151">O retorno de chamada `onCanceled` define a função que é executada quando a função é cancelada.</span><span class="sxs-lookup"><span data-stu-id="da76c-151">The `onCanceled` callback defines the function that executes when the function is canceled.</span></span>

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
> <span data-ttu-id="da76c-152">O Excel cancela a execução de uma função nas seguintes situações:</span><span class="sxs-lookup"><span data-stu-id="da76c-152">Excel cancels the execution of a function in the following situations:</span></span>
>
> - <span data-ttu-id="da76c-153">Quando o usuário edita ou exclui uma célula que faz referência à função.</span><span class="sxs-lookup"><span data-stu-id="da76c-153">When the user edits or deletes a cell that references the function.</span></span>
> - <span data-ttu-id="da76c-154">Quando é alterado um dos argumentos (entradas) para a função.</span><span class="sxs-lookup"><span data-stu-id="da76c-154">When one of the arguments (inputs) for the function changes.</span></span> <span data-ttu-id="da76c-155">Nesse caso, uma nova chamada de função é disparada, seguindo o cancelamento.</span><span class="sxs-lookup"><span data-stu-id="da76c-155">In this case, a new function call is triggered following the cancellation.</span></span>
> - <span data-ttu-id="da76c-156">Quando o usuário aciona manualmente um recálculo.</span><span class="sxs-lookup"><span data-stu-id="da76c-156">When the user triggers recalculation manually.</span></span> <span data-ttu-id="da76c-157">Nesse caso, uma nova chamada de função é disparada, seguindo o cancelamento.</span><span class="sxs-lookup"><span data-stu-id="da76c-157">In this case, a new function call is triggered following the cancellation.</span></span>

## <a name="next-steps"></a><span data-ttu-id="da76c-158">Próximas etapas</span><span class="sxs-lookup"><span data-stu-id="da76c-158">Next steps</span></span>

* <span data-ttu-id="da76c-159">Saiba mais sobre [diferentes tipos de parâmetros que as suas funções podem usar](custom-functions-parameter-options.md).</span><span class="sxs-lookup"><span data-stu-id="da76c-159">Learn about [different parameter types your functions can use](custom-functions-parameter-options.md).</span></span>
* <span data-ttu-id="da76c-160">Descubra como [agrupar várias chamadas de API](custom-functions-batching.md).</span><span class="sxs-lookup"><span data-stu-id="da76c-160">Discover how to [batch multiple API calls](custom-functions-batching.md).</span></span>

## <a name="see-also"></a><span data-ttu-id="da76c-161">Confira também</span><span class="sxs-lookup"><span data-stu-id="da76c-161">See also</span></span>

* [<span data-ttu-id="da76c-162">Valores voláteis nas funções</span><span class="sxs-lookup"><span data-stu-id="da76c-162">Volatile values in functions</span></span>](custom-functions-volatile.md)
* [<span data-ttu-id="da76c-163">Criar metadados JSON para funções personalizadas</span><span class="sxs-lookup"><span data-stu-id="da76c-163">Create JSON metadata for custom functions</span></span>](custom-functions-json-autogeneration.md)
* [<span data-ttu-id="da76c-164">Metadados de funções personalizadas</span><span class="sxs-lookup"><span data-stu-id="da76c-164">Custom functions metadata</span></span>](custom-functions-json.md)
* [<span data-ttu-id="da76c-165">Tempo de execução de funções personalizadas do Excel</span><span class="sxs-lookup"><span data-stu-id="da76c-165">Runtime for Excel custom functions</span></span>](custom-functions-runtime.md)
* <span data-ttu-id="da76c-166">[Práticas recomendadas de funções personalizadas](custom-functions-best-practices.md).</span><span class="sxs-lookup"><span data-stu-id="da76c-166">[Custom functions best practices](custom-functions-best-practices.md)</span></span>
* [<span data-ttu-id="da76c-167">Criar funções personalizadas no Excel</span><span class="sxs-lookup"><span data-stu-id="da76c-167">Create custom functions in Excel</span></span>](custom-functions-overview.md)
* [<span data-ttu-id="da76c-168">Tutorial de funções personalizadas do Excel</span><span class="sxs-lookup"><span data-stu-id="da76c-168">Excel custom functions tutorial</span></span>](../tutorials/excel-tutorial-create-custom-functions.md)
