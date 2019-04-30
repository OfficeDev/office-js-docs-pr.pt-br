---
ms.date: 04/20/2019
description: Solicite, transmita e cancele o fluxo de dados externos para sua pasta de trabalho com funções personalizadas no Excel
title: Solicitações da Web e outros dados de tratamento com funções personalizadas (prévia)
localization_priority: Priority
ms.openlocfilehash: 2942ec56e46d6eb586b516eedab17c1eeb98d9c8
ms.sourcegitcommit: 7462409209264dc7f8f89f3808a7a6249fcd739e
ms.translationtype: HT
ms.contentlocale: pt-BR
ms.lasthandoff: 04/26/2019
ms.locfileid: "33353262"
---
# <a name="receiving-and-handling-data-with-custom-functions"></a><span data-ttu-id="4630c-103">Recebimento e gerenciamento de dados com funções personalizadas</span><span class="sxs-lookup"><span data-stu-id="4630c-103">Receiving and handling data with custom functions</span></span>

<span data-ttu-id="4630c-104">Uma das maneiras pelas quais as funções personalizadas aprimoram o poder do Excel é receber dados de locais diferentes na pasta de trabalho, como a web ou um servidor (por meio de WebSockets).</span><span class="sxs-lookup"><span data-stu-id="4630c-104">One of the ways that custom functions enhance Excel's power is by receiving data from locations other than the workbook, such as the web or a server (through WebSockets).</span></span> <span data-ttu-id="4630c-105">As funções personalizadas podem solicitar dados por meio de XHR e buscar solicitações, bem como transmitir esses dados em tempo real.</span><span class="sxs-lookup"><span data-stu-id="4630c-105">Custom functions can request data through XHR and fetch requests as well as stream this data in real time.</span></span>

<span data-ttu-id="4630c-106">A documentação a seguir ilustra alguns exemplos de solicitações da web, mas para criar uma função de transmissão para você, experimente o [Tutorial de funções personalizadas](https://docs.microsoft.com/office/dev/add-ins/tutorials/excel-tutorial-create-custom-functions?tabs=excel-windows).</span><span class="sxs-lookup"><span data-stu-id="4630c-106">The documentation below illustrates some samples of web requests, but to build a streaming function for yourself, try the [Custom functions tutorial](https://docs.microsoft.com/office/dev/add-ins/tutorials/excel-tutorial-create-custom-functions?tabs=excel-windows).</span></span>

## <a name="functions-that-return-data-from-external-sources"></a><span data-ttu-id="4630c-107">Funções que retornam os dados de fontes externas</span><span class="sxs-lookup"><span data-stu-id="4630c-107">Functions that return data from external sources</span></span>

<span data-ttu-id="4630c-108">Se uma função personalizada recupera dados de uma fonte externa como na web, ela deve:</span><span class="sxs-lookup"><span data-stu-id="4630c-108">If a custom function retrieves data from an external source such as the web, it must:</span></span>

1. <span data-ttu-id="4630c-109">Retornar uma Promise do JavaScript para o Excel.</span><span class="sxs-lookup"><span data-stu-id="4630c-109">Return a JavaScript Promise to Excel.</span></span>
2. <span data-ttu-id="4630c-110">Resolva a promessa com o uso da função retorno de chamada de valor final.</span><span class="sxs-lookup"><span data-stu-id="4630c-110">Resolve the Promise with the final value using the callback function.</span></span>

<span data-ttu-id="4630c-111">É possível solicitar dados externos através de uma API como [ `Fetch` ](https://developer.mozilla.org/pt-BR/docs/Web/API/Fetch_API) ou usando `XmlHttpRequest` [(XHR)](https://developer.mozilla.org/pt-BR/docs/Web/API/XMLHttpRequest), uma API Web padrão que envia solicitações HTTP para interagir com os servidores.</span><span class="sxs-lookup"><span data-stu-id="4630c-111">You can request external data through an API like [`Fetch`](https://developer.mozilla.org/pt-BR/docs/Web/API/Fetch_API) or by using `XmlHttpRequest` [(XHR)](https://developer.mozilla.org/pt-BR/docs/Web/API/XMLHttpRequest), a standard web API that issues HTTP requests to interact with servers.</span></span>

<span data-ttu-id="4630c-112">No tempo de execução das funções personalizadas, o XHR implementa medidas de segurança adicionais solicitando uma [Política de mesma origem](https://developer.mozilla.org/pt-BR/docs/Web/Security/Same-origin_policy) ou um simples [CORS](https://www.w3.org/TR/cors/).</span><span class="sxs-lookup"><span data-stu-id="4630c-112">Within custom functions runtime, XHR implements additional security measures by requiring [Same Origin Policy](https://developer.mozilla.org/pt-BR/docs/Web/Security/Same-origin_policy) and simple [CORS](https://www.w3.org/TR/cors/).</span></span>

<span data-ttu-id="4630c-113">Observe que uma implementação CORS simples não pode usar cookies e é compatível apenas com métodos simples (GET, HEAD, POST).</span><span class="sxs-lookup"><span data-stu-id="4630c-113">Note that a simple CORS implementation cannot use cookies and only supports simple methods (GET, HEAD, POST).</span></span> <span data-ttu-id="4630c-114">A CORS simples aceita cabeçalhos simples com nomes de campos `Accept`, `Accept-Language`, `Content-Language`.</span><span class="sxs-lookup"><span data-stu-id="4630c-114">Simple CORS accepts simple headers with field names `Accept`, `Accept-Language`, `Content-Language`.</span></span> <span data-ttu-id="4630c-115">Você também pode usar um cabeçalho de tipo de conteúdo no CORS simples, desde que o tipo de conteúdo seja `application/x-www-form-urlencoded`, `text/plain`, ou `multipart/form-data`.</span><span class="sxs-lookup"><span data-stu-id="4630c-115">You can also use a Content-Type header in simple CORS, provided that the content type is `application/x-www-form-urlencoded`, `text/plain`, or `multipart/form-data`.</span></span>

### <a name="xhr-example"></a><span data-ttu-id="4630c-116">Exemplo de XHR</span><span class="sxs-lookup"><span data-stu-id="4630c-116">XHR example</span></span>

<span data-ttu-id="4630c-117">No código de exemplo a seguir, a função **getTemperature** chama a função sendWebRequest  para obter a temperatura de uma área específica, de acordo com a ID do termômetro.</span><span class="sxs-lookup"><span data-stu-id="4630c-117">In the following code sample, the **getTemperature** function calls the sendWebRequest function to get the temperature of a particular area based on thermometer ID.</span></span> <span data-ttu-id="4630c-118">A função sendWebRequest usa XHR para emitir uma solicitação GET para um ponto de extremidade que fornece os dados.</span><span class="sxs-lookup"><span data-stu-id="4630c-118">The sendWebRequest function uses XHR to issue a GET request to an endpoint that can provide the data.</span></span>

```JavaScript
function getTemperature(thermometerID) {
  return new Promise(function(setResult) {
      sendWebRequest(thermometerID, function(data){ 
          storeLastTemperature(thermometerID, data.temperature);
          setResult(data.temperature);
      });
  });
}

// Helper method that uses Office's implementation of XMLHttpRequest in the JavaScript runtime for custom functions  
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

<span data-ttu-id="4630c-119">Para outro exemplo de solicitação XHR com mais contexto, confira a função`getFile` dentro [deste arquivo](https://github.com/OfficeDev/Office-Add-in-JavaScript-FileDownload/blob/master/FileDownloadSampleWeb/Home.js) no repositório Github [Office-Add-in-JavaScript-FileDownload](https://github.com/OfficeDev/Office-Add-in-JavaScript-FileDownload).</span><span class="sxs-lookup"><span data-stu-id="4630c-119">For another sample of an XHR request with more context, see the `getFile` function within [this file](https://github.com/OfficeDev/Office-Add-in-JavaScript-FileDownload/blob/master/FileDownloadSampleWeb/Home.js) in the [Office-Add-in-JavaScript-FileDownload](https://github.com/OfficeDev/Office-Add-in-JavaScript-FileDownload) Github repository.</span></span>

### <a name="fetch-example"></a><span data-ttu-id="4630c-120">Exemplo de busca</span><span class="sxs-lookup"><span data-stu-id="4630c-120">Fetch example</span></span>

<span data-ttu-id="4630c-121">No seguinte exemplo de código, a função stockPriceStream usa um símbolo de cotação da bolsa para acessar o preço de uma ação a cada 1000 milissegundos.</span><span class="sxs-lookup"><span data-stu-id="4630c-121">In the following code sample, the stockPriceStream function uses a stock ticker symbol to get the price of a stock every 1000 milliseconds.</span></span> <span data-ttu-id="4630c-122">Para saber mais sobre este exemplo e obter as JSON acompanhante, confira a [tutorial de funções personalizados](https://docs.microsoft.com/office/dev/add-ins/tutorials/excel-tutorial-create-custom-functions?tabs=excel-windows#create-a-streaming-asynchronous-custom-function).</span><span class="sxs-lookup"><span data-stu-id="4630c-122">For more details about this sample and to get the accompanying JSON, see the [Custom functions tutorial](https://docs.microsoft.com/office/dev/add-ins/tutorials/excel-tutorial-create-custom-functions?tabs=excel-windows#create-a-streaming-asynchronous-custom-function).</span></span> 

```JavaScript
function stockPriceStream(ticker, handler) {
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
                handler.setResult(parseFloat(text));
            })
            .catch(function(error) {
                handler.setResult(error);
            })
            .then(function() {
                isPending = false;
            });
    }, updateFrequency);

    handler.onCanceled = () => {
        clearInterval(timer);
    };
}

CustomFunctions.associate("STOCKPRICESTREAM", stockPriceStream);
```

## <a name="receiving-data-via-websockets"></a><span data-ttu-id="4630c-123">Como receber dados por meio de WebSockets</span><span class="sxs-lookup"><span data-stu-id="4630c-123">Receiving data via WebSockets</span></span>

<span data-ttu-id="4630c-124">Em uma função personalizada, é possível usar WebSockets para trocar dados por meio de uma conexão persistente com um servidor.</span><span class="sxs-lookup"><span data-stu-id="4630c-124">Within a custom function, you can use WebSockets to exchange data over a persistent connection with a server.</span></span> <span data-ttu-id="4630c-125">Usando WebSockets, a função personalizada pode abrir uma conexão com um servidor e, em seguida, receber mensagens do servidor automaticamente, quando determinados eventos ocorrerem, sem precisar consultar explicitamente os dados do servidor.</span><span class="sxs-lookup"><span data-stu-id="4630c-125">By using WebSockets, your custom function can open a connection with a server and then automatically receive messages from the server when certain events occur, without having to explicitly poll the server for data.</span></span>

### <a name="websockets-example"></a><span data-ttu-id="4630c-126">Exemplo de WebSockets</span><span class="sxs-lookup"><span data-stu-id="4630c-126">WebSockets example</span></span>

<span data-ttu-id="4630c-127">O código de exemplo a seguir estabelece uma conexão WebSocket e registra cada mensagem de entrada do servidor.</span><span class="sxs-lookup"><span data-stu-id="4630c-127">The following code sample establishes a WebSocket connection and then logs each incoming message from the server.</span></span>

```JavaScript
var ws = new WebSocket('wss://bundles.office.com');

ws.onmessage(message) {
    console.log(`Recieved: ${message}`);
}

ws.onerror(error){
    console.err(`Failed: ${error}`);
}
```

## <a name="streaming-functions"></a><span data-ttu-id="4630c-128">Funções Streaming</span><span class="sxs-lookup"><span data-stu-id="4630c-128">Streaming functions</span></span>

<span data-ttu-id="4630c-129">Funções personalizadas de streaming permitem a saída de dados das células repetidamente ao longo do tempo, sem a necessidade de um usuário explicitamente solicitar a atualização de dados.</span><span class="sxs-lookup"><span data-stu-id="4630c-129">Streaming custom functions enable you to output data to cells repeatedly over time, without requiring a user to explicitly request data refresh.</span></span> <span data-ttu-id="4630c-130">O exemplo a seguir é uma função personalizada que adiciona um número ao resultado a cada segundo.</span><span class="sxs-lookup"><span data-stu-id="4630c-130">The following code sample is a custom function that adds a number to the result every second.</span></span> <span data-ttu-id="4630c-131">Observe o seguinte sobre este código:</span><span class="sxs-lookup"><span data-stu-id="4630c-131">Note the following about this code:</span></span>

- <span data-ttu-id="4630c-132">Cada novo valor usando o Excel automaticamente exibirá o retorno de chamada setResult.</span><span class="sxs-lookup"><span data-stu-id="4630c-132">Excel displays each new value automatically using the setResult callback.</span></span>
- <span data-ttu-id="4630c-133">O segundo parâmetro de entrada, identificador, não é exibido para os usuários finais no Excel quando eles selecionam a função no menu de preenchimento automático.</span><span class="sxs-lookup"><span data-stu-id="4630c-133">The second input parameter, handler, is not displayed to end users in Excel when they select the function from the autocomplete menu.</span></span>
- <span data-ttu-id="4630c-134">O retorno de chamada onCanceled define a função que é executada quando a função é cancelada.</span><span class="sxs-lookup"><span data-stu-id="4630c-134">The onCanceled callback defines the function that executes when the function is canceled.</span></span> <span data-ttu-id="4630c-135">Implemente um identificador de cancelamento assim para qualquer função de streaming.</span><span class="sxs-lookup"><span data-stu-id="4630c-135">You must implement a cancellation handler like this for any streaming function.</span></span> <span data-ttu-id="4630c-136">Para saber mais, confira [Cancelar uma função](#canceling-a-function).</span><span class="sxs-lookup"><span data-stu-id="4630c-136">For more information, see [Canceling a function](#canceling-a-function).</span></span>

```JavaScript
function incrementValue(increment, handler){
  var result = 0;
  setInterval(function(){
    result += increment;
    handler.setResult(result);
  }, 1000);

  handler.onCanceled = function(){
    clearInterval(timer);
  }
}

CustomFunctions.associate("INCREMENTVALUE", incrementValue);
```

<span data-ttu-id="4630c-137">Quando você especifica metadados para uma função de streaming no arquivo de metadados JSON, é possível gerar isso automaticamente usando uma tag `@streaming` de comentário JSDOC no arquivo de script da sua função.</span><span class="sxs-lookup"><span data-stu-id="4630c-137">When you specify metadata for a streaming function in the JSON metadata file, you can autogenerate this by using a `@streaming` JSDOC comment tag in your function's script file.</span></span> <span data-ttu-id="4630c-138">Para mais detalhes, consulte [Criar metadados JSON para funções personalizadas](custom-functions-json-autogeneration.md).</span><span class="sxs-lookup"><span data-stu-id="4630c-138">For more details, see [Create JSON metadata for custom functions](custom-functions-json-autogeneration.md).</span></span>

## <a name="canceling-a-function"></a><span data-ttu-id="4630c-139">Cancelar uma função</span><span class="sxs-lookup"><span data-stu-id="4630c-139">Canceling a function</span></span>

<span data-ttu-id="4630c-140">Em algumas situações, talvez seja necessário cancelar a execução de uma função personalizada de streaming para reduzir o consumo de banda larga, memória de trabalho e carregamento de CPU.</span><span class="sxs-lookup"><span data-stu-id="4630c-140">In some situations, you may need to cancel the execution of a streaming custom function to reduce its bandwidth consumption, working memory, and CPU load.</span></span> <span data-ttu-id="4630c-141">O Excel cancela a execução de uma função nas seguintes situações:</span><span class="sxs-lookup"><span data-stu-id="4630c-141">Excel cancels the execution of a function in the following situations:</span></span>

- <span data-ttu-id="4630c-142">Quando o usuário edita ou exclui uma célula que faz referência à função.</span><span class="sxs-lookup"><span data-stu-id="4630c-142">When the user edits or deletes a cell that references the function.</span></span>
- <span data-ttu-id="4630c-143">Quando é alterado um dos argumentos (entradas) para a função.</span><span class="sxs-lookup"><span data-stu-id="4630c-143">When one of the arguments (inputs) for the function changes.</span></span> <span data-ttu-id="4630c-144">Nesse caso, uma nova chamada de função é disparada, seguindo o cancelamento.</span><span class="sxs-lookup"><span data-stu-id="4630c-144">In this case, a new function call is triggered following the cancellation.</span></span>
- <span data-ttu-id="4630c-145">Quando o usuário aciona manualmente um recálculo.</span><span class="sxs-lookup"><span data-stu-id="4630c-145">When the user triggers recalculation manually.</span></span> <span data-ttu-id="4630c-146">Nesse caso, uma nova chamada de função é disparada, seguindo o cancelamento.</span><span class="sxs-lookup"><span data-stu-id="4630c-146">In this case, a new function call is triggered following the cancellation.</span></span>

<span data-ttu-id="4630c-147">Para tornar uma função possível de ser cancelada, implemente um identificador de código de função para informar o que fazer quando ela for cancelada.</span><span class="sxs-lookup"><span data-stu-id="4630c-147">To make a function cancelable, implement a handler in your function's code to tell it what to do when it is canceled.</span></span> <span data-ttu-id="4630c-148">Além disso, use a tag `@cancelable` de comentário JSDOC no arquivo de script da sua função.</span><span class="sxs-lookup"><span data-stu-id="4630c-148">Additionally, use the `@cancelable` JSDOC comment tag in your function's script file.</span></span> <span data-ttu-id="4630c-149">Para mais detalhes, consulte [Criar metadados JSON para funções personalizadas](custom-functions-json-autogeneration.md).</span><span class="sxs-lookup"><span data-stu-id="4630c-149">For more details, see [Create JSON metadata for custom functions](custom-functions-json-autogeneration.md).</span></span>

## <a name="see-also"></a><span data-ttu-id="4630c-150">Confira também</span><span class="sxs-lookup"><span data-stu-id="4630c-150">See also</span></span>

* [<span data-ttu-id="4630c-151">Tutorial de funções personalizadas do Excel</span><span class="sxs-lookup"><span data-stu-id="4630c-151">Excel custom functions tutorial</span></span>](../tutorials/excel-tutorial-create-custom-functions.md)
* [<span data-ttu-id="4630c-152">Metadados de funções personalizadas</span><span class="sxs-lookup"><span data-stu-id="4630c-152">Custom functions metadata</span></span>](custom-functions-json.md)
* [<span data-ttu-id="4630c-153">Criar metadados JSON para funções personalizadas</span><span class="sxs-lookup"><span data-stu-id="4630c-153">Create JSON metadata for custom functions (preview)</span></span>](custom-functions-json-autogeneration.md)
* [<span data-ttu-id="4630c-154">Tempo de execução de funções personalizadas do Excel</span><span class="sxs-lookup"><span data-stu-id="4630c-154">Runtime for Excel custom functions</span></span>](custom-functions-runtime.md)
* <span data-ttu-id="4630c-155">[Práticas recomendadas de funções personalizadas](custom-functions-best-practices.md).</span><span class="sxs-lookup"><span data-stu-id="4630c-155">[Custom functions best practices](custom-functions-best-practices.md)</span></span>
* [<span data-ttu-id="4630c-156">Log de alteração de funções personalizadas</span><span class="sxs-lookup"><span data-stu-id="4630c-156">Custom functions changelog</span></span>](custom-functions-changelog.md)
