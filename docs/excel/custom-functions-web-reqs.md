---
ms.date: 01/14/2020
description: Solicite, transmita e cancele o fluxo de dados externos para sua pasta de trabalho com funções personalizadas no Excel
title: Receber e tratar dados com funções personalizadas
localization_priority: Normal
ms.openlocfilehash: c59133389311721c4d0507337a17b346af97e690
ms.sourcegitcommit: d15bca2c12732f8599be2ec4b2adc7c254552f52
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 02/12/2020
ms.locfileid: "41949703"
---
# <a name="receive-and-handle-data-with-custom-functions"></a><span data-ttu-id="c8e00-103">Receber e tratar dados com funções personalizadas</span><span class="sxs-lookup"><span data-stu-id="c8e00-103">Receive and handle data with custom functions</span></span>

<span data-ttu-id="c8e00-104">Uma das maneiras pelas quais as funções personalizadas aprimoram o poder do Excel é através do recebimento de dados de outros locais diferente da pasta de trabalho, como a Web ou um servidor (por meio de WebSockets).</span><span class="sxs-lookup"><span data-stu-id="c8e00-104">One of the ways that custom functions enhances Excel's power is by receiving data from locations other than the workbook, such as the web or a server (through WebSockets).</span></span> <span data-ttu-id="c8e00-105">É possível solicitar dados externos através de uma API como [ `Fetch` ](https://developer.mozilla.org/en-US/docs/Web/API/Fetch_API) ou usando `XmlHttpRequest` [(XHR)](https://developer.mozilla.org/en-US/docs/Web/API/XMLHttpRequest), uma API Web padrão que envia solicitações HTTP para interagir com os servidores.</span><span class="sxs-lookup"><span data-stu-id="c8e00-105">You can request external data through an API like [`Fetch`](https://developer.mozilla.org/en-US/docs/Web/API/Fetch_API) or by using `XmlHttpRequest` [(XHR)](https://developer.mozilla.org/en-US/docs/Web/API/XMLHttpRequest), a standard web API that issues HTTP requests to interact with servers.</span></span>

[!include[Excel custom functions note](../includes/excel-custom-functions-note.md)]

![Um gif de uma função personalizada que transmite o tempo de uma API](../images/custom-functions-web-api.gif)

## <a name="functions-that-return-data-from-external-sources"></a><span data-ttu-id="c8e00-107">Funções que retornam os dados de fontes externas</span><span class="sxs-lookup"><span data-stu-id="c8e00-107">Functions that return data from external sources</span></span>

<span data-ttu-id="c8e00-108">Se uma função personalizada recupera dados de uma fonte externa como na web, ela deve:</span><span class="sxs-lookup"><span data-stu-id="c8e00-108">If a custom function retrieves data from an external source such as the web, it must:</span></span>

1. <span data-ttu-id="c8e00-109">Retornar uma Promise do JavaScript para o Excel.</span><span class="sxs-lookup"><span data-stu-id="c8e00-109">Return a JavaScript Promise to Excel.</span></span>
2. <span data-ttu-id="c8e00-110">Resolva a promessa com o uso da função retorno de chamada de valor final.</span><span class="sxs-lookup"><span data-stu-id="c8e00-110">Resolve the Promise with the final value using the callback function.</span></span>

### <a name="fetch-example"></a><span data-ttu-id="c8e00-111">Exemplo de busca</span><span class="sxs-lookup"><span data-stu-id="c8e00-111">Fetch example</span></span>

<span data-ttu-id="c8e00-112">No exemplo a seguir, a função \*\* WebRequest\*\* chega à API hipotética "número de pessoas no espaço", que controla o número de pessoas que estão na estação de espaço internacional.</span><span class="sxs-lookup"><span data-stu-id="c8e00-112">In the following code sample, the **webRequest** function reaches out to the hypothetical Contoso "Number of People in Space" API, which tracks the number of people currently on the International Space Station.</span></span> <span data-ttu-id="c8e00-113">A função retorna uma promessa de JavaScript e usa fetch para solicitar informações da API.</span><span class="sxs-lookup"><span data-stu-id="c8e00-113">The function returns a JavaScript Promise and uses fetch to request information from the API.</span></span> <span data-ttu-id="c8e00-114">Os dados resultantes são transformados em JSON e a`names` propriedade é convertida em uma cadeia de caracteres, que é usada para resolver a promessa.</span><span class="sxs-lookup"><span data-stu-id="c8e00-114">The resulting data is transformed into JSON and the `names` property is converted into a string, which is used to resolve the Promise.</span></span>

<span data-ttu-id="c8e00-115">Ao desenvolver suas próprias funções, talvez você queira executar uma ação caso a solicitação da Web não tenha sido concluída de maneira oportuna ou considere [o envio de várias solicitações](./custom-functions-batching.md)da API.</span><span class="sxs-lookup"><span data-stu-id="c8e00-115">When developing your own functions, you may want to perform an action if the web request does not complete in a timely manner or consider [batching up multiple API requests](./custom-functions-batching.md).</span></span>

```JS
/**
 * Requests the names of the people currently on the International Space Station from a hypothetical API.
 * @customfunction
 */
function webRequest() {
  let url = "https://www.contoso.com/NumberOfPeopleInSpace";
  return new Promise(function (resolve, reject) {
    fetch(url)
      .then(function (response){
        return response.json();
        }
      )
      .then(function (json) {
        resolve(JSON.stringify(json.names));
      })
  })
}
```

>[!NOTE]
><span data-ttu-id="c8e00-116">Usar `Fetch` evita retornos de chamada aninhados e pode ser preferível do XHR em alguns casos.</span><span class="sxs-lookup"><span data-stu-id="c8e00-116">Using `Fetch` avoids nested callbacks and may be preferable to XHR in some cases.</span></span>

### <a name="xhr-example"></a><span data-ttu-id="c8e00-117">Exemplo de XHR</span><span class="sxs-lookup"><span data-stu-id="c8e00-117">XHR example</span></span>

<span data-ttu-id="c8e00-118">No tempo de execução das funções personalizadas, o XHR implementa medidas de segurança adicionais solicitando uma [Política de mesma origem](https://developer.mozilla.org/en-US/docs/Web/Security/Same-origin_policy) ou um simples [CORS](https://www.w3.org/TR/cors/).</span><span class="sxs-lookup"><span data-stu-id="c8e00-118">Within custom functions runtime, XHR implements additional security measures by requiring [Same Origin Policy](https://developer.mozilla.org/en-US/docs/Web/Security/Same-origin_policy) and simple [CORS](https://www.w3.org/TR/cors/).</span></span>

<span data-ttu-id="c8e00-119">Observe que uma implementação CORS simples não pode usar cookies e é compatível apenas com métodos simples (GET, HEAD, POST).</span><span class="sxs-lookup"><span data-stu-id="c8e00-119">Note that a simple CORS implementation cannot use cookies and only supports simple methods (GET, HEAD, POST).</span></span> <span data-ttu-id="c8e00-120">A CORS simples aceita cabeçalhos simples com nomes de campos `Accept`, `Accept-Language`, `Content-Language`.</span><span class="sxs-lookup"><span data-stu-id="c8e00-120">Simple CORS accepts simple headers with field names `Accept`, `Accept-Language`, `Content-Language`.</span></span> <span data-ttu-id="c8e00-121">Você também pode usar um cabeçalho de tipo de conteúdo no CORS simples, desde que o tipo de conteúdo seja `application/x-www-form-urlencoded`, `text/plain`, ou `multipart/form-data`.</span><span class="sxs-lookup"><span data-stu-id="c8e00-121">You can also use a Content-Type header in simple CORS, provided that the content type is `application/x-www-form-urlencoded`, `text/plain`, or `multipart/form-data`.</span></span>

<span data-ttu-id="c8e00-122">No exemplo a seguir, a função\*\* getStarCount\*\* chama o GitHub API para descobrir o valor de estrelas fornecido para o repositório de um determinado usuário.</span><span class="sxs-lookup"><span data-stu-id="c8e00-122">In the following code sample, the **getStarCount** function calls the Github API to discover the amount of stars given to a particular user's repository.</span></span> <span data-ttu-id="c8e00-123">Essa é uma função assíncrona que retorna uma promessa de JavaScript.</span><span class="sxs-lookup"><span data-stu-id="c8e00-123">This is an asynchronous function which returns a JavaScript Promise.</span></span> <span data-ttu-id="c8e00-124">Quando os dados forem obtidos da chamada da Web, a promessa será resolvida, que retornará os dados para a célula.</span><span class="sxs-lookup"><span data-stu-id="c8e00-124">When data is obtained from the web call, the Promise is resolved which returns the data to the cell.</span></span>

```TS
/**
 * Gets the star count for a given Github organization or user and repository.
 * @customfunction
 * @param userName string name of organization or user.
 * @param repoName string name of the repository.
 * @return number of stars.
 */

async function getStarCount(userName: string, repoName: string) {

  const url = "https://api.github.com/repos/" + userName + "/" + repoName;

  let xhttp = new XMLHttpRequest();

  return new Promise(function(resolve, reject) {
    xhttp.onreadystatechange = function() {
      if (xhttp.readyState !== 4) return;

      if (xhttp.status == 200) {
        resolve(JSON.parse(xhttp.responseText).watchers_count);
      } else {
        reject({
          status: xhttp.status,

          statusText: xhttp.statusText
        });
      }
    };

    xhttp.open("GET", url, true);

    xhttp.send();
  });
}
```

## <a name="make-a-streaming-function"></a><span data-ttu-id="c8e00-125">Faça uma função de streaming</span><span class="sxs-lookup"><span data-stu-id="c8e00-125">Make a streaming function</span></span>

<span data-ttu-id="c8e00-126">Funções personalizadas de streaming permitem a saída de dados para células que atualizam repetidamente, sem a necessidade de um usuário explicitamente atualizar coisa alguma.</span><span class="sxs-lookup"><span data-stu-id="c8e00-126">Streaming custom functions enable you to output data to cells that updates repeatedly, without requiring a user to explicitly refresh anything.</span></span> <span data-ttu-id="c8e00-127">Isso pode ser útil para verificar dados ativos de um serviço online, como a função no [tutorial de funções personalizadas](../tutorials/excel-tutorial-create-custom-functions.md).</span><span class="sxs-lookup"><span data-stu-id="c8e00-127">This can be useful to check live data from a service online, like the function in [the custom functions tutorial](../tutorials/excel-tutorial-create-custom-functions.md).</span></span>

<span data-ttu-id="c8e00-128">Para declarar uma função de streaming, use a marca `@streaming` ou use o parâmetro `CustomFunctions.StreamingInvocation` de invocação, que indicará que a função é streaming.</span><span class="sxs-lookup"><span data-stu-id="c8e00-128">To declare a streaming function, either use the tag `@streaming` or make use of the `CustomFunctions.StreamingInvocation` invocation parameter, which will indicate that your function is streaming.</span></span> <span data-ttu-id="c8e00-129">Para alertar os usuários para o fato de que sua função pode ser reavaliada com base em novas informações, considere colocar fluxo ou outro texto para indicar isso no nome ou na descrição de sua função.</span><span class="sxs-lookup"><span data-stu-id="c8e00-129">To alert users to the fact that your function may re-evaluate based on new information, consider putting stream or other wording to indicate this in the name or description of your function.</span></span>

<span data-ttu-id="c8e00-130">O exemplo a seguir é uma função personalizada que adiciona um número ao resultado a cada segundo.</span><span class="sxs-lookup"><span data-stu-id="c8e00-130">The following code sample is a custom function that adds a number to the result every second.</span></span> <span data-ttu-id="c8e00-131">Observe o seguinte sobre este código:</span><span class="sxs-lookup"><span data-stu-id="c8e00-131">Note the following about this code:</span></span>

- <span data-ttu-id="c8e00-132">O Excel exibe cada valor novo automaticamente usando o método `setResult`.</span><span class="sxs-lookup"><span data-stu-id="c8e00-132">Excel displays each new value automatically using the `setResult` method.</span></span>
- <span data-ttu-id="c8e00-133">O segundo parâmetro de entrada, invocação, não é exibido para os usuários finais no Excel quando eles selecionam a função no menu de preenchimento automático.</span><span class="sxs-lookup"><span data-stu-id="c8e00-133">The second input parameter, invocation, is not displayed to end users in Excel when they select the function from the autocomplete menu.</span></span>
- <span data-ttu-id="c8e00-134">O retorno de chamada `onCanceled` define a função que é executada quando a função é cancelada.</span><span class="sxs-lookup"><span data-stu-id="c8e00-134">The `onCanceled` callback defines the function that executes when the function is canceled.</span></span>
- <span data-ttu-id="c8e00-135">O streaming não está necessariamente vinculado a fazer uma solicitação na Web: nesse caso, a função não está fazendo uma solicitação da Web, mas ainda está com dados em intervalos definidos, portanto, é `invocation` necessário usar o parâmetro de streaming.</span><span class="sxs-lookup"><span data-stu-id="c8e00-135">Streaming isn't necessarily tied to making a web request: in this case, the function isn't making a web request but is still getting data at set intervals, so it requires the use of the streaming `invocation` parameter.</span></span>

```js
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
```

<span data-ttu-id="c8e00-136">Além de conhecer o `onCanceled` retorno de chamada, você também deve saber se o Excel cancela a execução de uma função nas seguintes situações:</span><span class="sxs-lookup"><span data-stu-id="c8e00-136">In addition to knowing about the `onCanceled` callback, you should also know that Excel cancels the execution of a function in the following situations:</span></span>

- <span data-ttu-id="c8e00-137">Quando o usuário edita ou exclui uma célula que faz referência à função.</span><span class="sxs-lookup"><span data-stu-id="c8e00-137">When the user edits or deletes a cell that references the function.</span></span>
- <span data-ttu-id="c8e00-138">Quando é alterado um dos argumentos (entradas) para a função.</span><span class="sxs-lookup"><span data-stu-id="c8e00-138">When one of the arguments (inputs) for the function changes.</span></span> <span data-ttu-id="c8e00-139">Nesse caso, uma nova chamada de função é disparada, seguindo o cancelamento.</span><span class="sxs-lookup"><span data-stu-id="c8e00-139">In this case, a new function call is triggered following the cancellation.</span></span>
- <span data-ttu-id="c8e00-140">Quando o usuário aciona manualmente um recálculo.</span><span class="sxs-lookup"><span data-stu-id="c8e00-140">When the user triggers recalculation manually.</span></span> <span data-ttu-id="c8e00-141">Nesse caso, uma nova chamada de função é disparada, seguindo o cancelamento.</span><span class="sxs-lookup"><span data-stu-id="c8e00-141">In this case, a new function call is triggered following the cancellation.</span></span>

<span data-ttu-id="c8e00-142">Você também pode considerar a definição de um valor de streaming padrão para lidar com os casos em que uma solicitação for feita, mas você está offline.</span><span class="sxs-lookup"><span data-stu-id="c8e00-142">You can also consider setting a default streaming value to handle cases when a request is made but you are offline.</span></span>

> [!NOTE]
> <span data-ttu-id="c8e00-143">Observe que há também uma categoria de funções chamada de funções canceláveis, que _não_ estão relacionadas a funções de streaming.</span><span class="sxs-lookup"><span data-stu-id="c8e00-143">Note that there are also a category of functions called cancelable functions, which are _not_ related to streaming functions.</span></span> <span data-ttu-id="c8e00-144">Versões anteriores de funções personalizadas exigiam que você declarasse `"cancelable": true` e `"streaming": true` em JSON escrito à mão.</span><span class="sxs-lookup"><span data-stu-id="c8e00-144">Previous versions of custom functions required you to declare `"cancelable": true` and `"streaming": true` in JSON written by hand.</span></span> <span data-ttu-id="c8e00-145">Desde a introdução de metadados autogerados, somente as funções personalizadas assíncronas que retornam um único valor são canceláveis.</span><span class="sxs-lookup"><span data-stu-id="c8e00-145">Since the introduction of autogenerated metadata, only asynchronous custom functions which return one value are cancelable.</span></span> <span data-ttu-id="c8e00-146">Funções canceláveis permitem que uma solicitação da Web seja encerrada no meio de uma solicitação, usando um [`CancelableInvocation`](/javascript/api/custom-functions-runtime/customfunctions.cancelableinvocation) para decidir o que fazer após o cancelamento.</span><span class="sxs-lookup"><span data-stu-id="c8e00-146">Cancelable functions allow a web request to be terminated in the middle of a request, using a [`CancelableInvocation`](/javascript/api/custom-functions-runtime/customfunctions.cancelableinvocation) to decide what to do upon cancellation.</span></span> <span data-ttu-id="c8e00-147">Declare uma função cancelável usando a tag `@cancelable`.</span><span class="sxs-lookup"><span data-stu-id="c8e00-147">Declare a cancelable function using the tag `@cancelable`.</span></span>

### <a name="using-an-invocation-parameter"></a><span data-ttu-id="c8e00-148">Usando um parâmetro de invocação</span><span class="sxs-lookup"><span data-stu-id="c8e00-148">Using an invocation parameter</span></span>

<span data-ttu-id="c8e00-149">O parâmetro `invocation` é o último parâmetro de qualquer função personalizada por padrão.</span><span class="sxs-lookup"><span data-stu-id="c8e00-149">The `invocation` parameter is the last parameter of any custom function by default.</span></span> <span data-ttu-id="c8e00-150">O parâmetro `invocation` fornece um contexto sobre a célula (como o seu endereço) e também permite com que você use os métodos `setResult` e `onCanceled`.</span><span class="sxs-lookup"><span data-stu-id="c8e00-150">The `invocation` parameter gives context about the cell (such as its address and contents) and also allows you to use `setResult` and `onCanceled` methods.</span></span> <span data-ttu-id="c8e00-151">Esses métodos definem o que uma função faz quando a função transmite (`setResult`) ou é cancelada (`onCanceled`).</span><span class="sxs-lookup"><span data-stu-id="c8e00-151">These methods define what a function does when the function streams (`setResult`) or is canceled (`onCanceled`).</span></span>

<span data-ttu-id="c8e00-152">Se você estiver usando o TypeScript, o manipulador de invocações deve ser do tipo `CustomFunctions.StreamingInvocation` ou `CustomFunctions.CancelableInvocation`.</span><span class="sxs-lookup"><span data-stu-id="c8e00-152">If you're using TypeScript, the invocation handler needs to be of type `CustomFunctions.StreamingInvocation` or `CustomFunctions.CancelableInvocation`.</span></span>

## <a name="receive-data-via-websockets"></a><span data-ttu-id="c8e00-153">Receber dados por meio de WebSockets</span><span class="sxs-lookup"><span data-stu-id="c8e00-153">Receive data via WebSockets</span></span>

<span data-ttu-id="c8e00-154">Em uma função personalizada, é possível usar WebSockets para trocar dados por meio de uma conexão persistente com um servidor.</span><span class="sxs-lookup"><span data-stu-id="c8e00-154">Within a custom function, you can use WebSockets to exchange data over a persistent connection with a server.</span></span> <span data-ttu-id="c8e00-155">Usando WebSockets, a função personalizada pode abrir uma conexão com um servidor e, em seguida, receber mensagens do servidor automaticamente, quando determinados eventos ocorrerem, sem precisar consultar explicitamente os dados do servidor.</span><span class="sxs-lookup"><span data-stu-id="c8e00-155">By using WebSockets, your custom function can open a connection with a server and then automatically receive messages from the server when certain events occur, without having to explicitly poll the server for data.</span></span>

### <a name="websockets-example"></a><span data-ttu-id="c8e00-156">Exemplo de WebSockets</span><span class="sxs-lookup"><span data-stu-id="c8e00-156">WebSockets example</span></span>

<span data-ttu-id="c8e00-157">O código de exemplo a seguir estabelece uma conexão WebSocket e registra cada mensagem de entrada do servidor.</span><span class="sxs-lookup"><span data-stu-id="c8e00-157">The following code sample establishes a WebSocket connection and then logs each incoming message from the server.</span></span>

```js
let ws = new WebSocket('wss://bundles.office.com');

ws.onmessage(message) {
    console.log(`Received: ${message}`);
}

ws.onerror(error){
    console.err(`Failed: ${error}`);
}
```

## <a name="next-steps"></a><span data-ttu-id="c8e00-158">Próximas etapas</span><span class="sxs-lookup"><span data-stu-id="c8e00-158">Next steps</span></span>

- <span data-ttu-id="c8e00-159">Saiba mais sobre [diferentes tipos de parâmetros que as suas funções podem usar](custom-functions-parameter-options.md).</span><span class="sxs-lookup"><span data-stu-id="c8e00-159">Learn about [different parameter types your functions can use](custom-functions-parameter-options.md).</span></span>
- <span data-ttu-id="c8e00-160">Descubra como [agrupar várias chamadas de API](custom-functions-batching.md).</span><span class="sxs-lookup"><span data-stu-id="c8e00-160">Discover how to [batch multiple API calls](custom-functions-batching.md).</span></span>

## <a name="see-also"></a><span data-ttu-id="c8e00-161">Confira também</span><span class="sxs-lookup"><span data-stu-id="c8e00-161">See also</span></span>

- [<span data-ttu-id="c8e00-162">Valores voláteis nas funções</span><span class="sxs-lookup"><span data-stu-id="c8e00-162">Volatile values in functions</span></span>](custom-functions-volatile.md)
- [<span data-ttu-id="c8e00-163">Criar metadados JSON para funções personalizadas</span><span class="sxs-lookup"><span data-stu-id="c8e00-163">Create JSON metadata for custom functions</span></span>](custom-functions-json-autogeneration.md)
- [<span data-ttu-id="c8e00-164">Metadados de funções personalizadas</span><span class="sxs-lookup"><span data-stu-id="c8e00-164">Custom functions metadata</span></span>](custom-functions-json.md)
- [<span data-ttu-id="c8e00-165">Tempo de execução de funções personalizadas do Excel</span><span class="sxs-lookup"><span data-stu-id="c8e00-165">Runtime for Excel custom functions</span></span>](custom-functions-runtime.md)
- [<span data-ttu-id="c8e00-166">Criar funções personalizadas no Excel</span><span class="sxs-lookup"><span data-stu-id="c8e00-166">Create custom functions in Excel</span></span>](custom-functions-overview.md)
- [<span data-ttu-id="c8e00-167">Tutorial de funções personalizadas do Excel</span><span class="sxs-lookup"><span data-stu-id="c8e00-167">Excel custom functions tutorial</span></span>](../tutorials/excel-tutorial-create-custom-functions.md)
