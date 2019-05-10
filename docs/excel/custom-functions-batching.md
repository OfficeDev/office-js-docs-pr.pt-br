---
ms.date: 05/03/2019
description: Reúna as funções personalizadas em lotes para reduzir as chamadas de rede para um serviço remoto.
title: Enviando em lote chamadas de função personalizada para um serviço remoto
localization_priority: Priority
ms.openlocfilehash: da9f3ee3243b52df5d49f32c8ab6cbada97e17ca
ms.sourcegitcommit: ff73cc04e5718765fcbe74181505a974db69c3f5
ms.translationtype: HT
ms.contentlocale: pt-BR
ms.lasthandoff: 05/06/2019
ms.locfileid: "33628127"
---
# <a name="batching-custom-function-calls-for-a-remote-service"></a><span data-ttu-id="0e4bf-103">Enviando em lote chamadas de função personalizada para um serviço remoto</span><span class="sxs-lookup"><span data-stu-id="0e4bf-103">Batching custom function calls for a remote service</span></span>

<span data-ttu-id="0e4bf-104">Se as suas funções personalizadas chamarem um serviço remoto, você poderá usar um padrão de envio em lotes para reduzir o número de chamadas de rede para o serviço remoto.</span><span class="sxs-lookup"><span data-stu-id="0e4bf-104">If your custom functions call a remote service you can use a batching pattern to reduce the number of network calls to the remote service.</span></span> <span data-ttu-id="0e4bf-105">Para reduzir a idas e voltas na rede, você reúne todas as chamadas em uma única chamada para o serviço da Web.</span><span class="sxs-lookup"><span data-stu-id="0e4bf-105">To reduce network round trips you batch all the calls into a single call to the web service.</span></span> <span data-ttu-id="0e4bf-106">Isso é ideal quando a planilha é recalculada.</span><span class="sxs-lookup"><span data-stu-id="0e4bf-106">This is ideal when the spreadsheet is recalculated.</span></span>

<span data-ttu-id="0e4bf-107">Por exemplo, se alguém usou sua função personalizada em 100 células em uma planilha e depois recalculou a planilha, sua função personalizada seria executada 100 vezes e faria 100 chamadas de rede.</span><span class="sxs-lookup"><span data-stu-id="0e4bf-107">For example, if someone used your custom function in 100 cells in a spreadsheet, and then recalculated the spreadsheet, your custom function would run 100 times and make 100 network calls.</span></span> <span data-ttu-id="0e4bf-108">Usando um padrão de envio em lotes, as chamadas podem ser combinadas para fazer todos os 100 cálculos em uma única chamada de rede.</span><span class="sxs-lookup"><span data-stu-id="0e4bf-108">By using a batching pattern, the calls can be combined to make all 100 calculations in a single network call.</span></span>

[!include[Excel custom functions note](../includes/excel-custom-functions-note.md)]

## <a name="view-the-completed-sample"></a><span data-ttu-id="0e4bf-109">Ver o exemplo concluído</span><span class="sxs-lookup"><span data-stu-id="0e4bf-109">View the completed sample</span></span>

<span data-ttu-id="0e4bf-110">Você pode seguir este artigo e colar os exemplos de código em seu próprio projeto.</span><span class="sxs-lookup"><span data-stu-id="0e4bf-110">You can follow this article and paste the code examples into your own project.</span></span> <span data-ttu-id="0e4bf-111">Por exemplo, você pode usar o [gerador do Yo Office](https://github.com/OfficeDev/generator-office) para criar um novo projeto de função personalizada para TypeScript e, em seguida, adicionar todo o código deste artigo ao projeto.</span><span class="sxs-lookup"><span data-stu-id="0e4bf-111">For example, you can use yo office to create a new custom function project for TypeScript, then add all the code from this article to the project.</span></span> <span data-ttu-id="0e4bf-112">Você pode então executar o código e experimentá-lo.</span><span class="sxs-lookup"><span data-stu-id="0e4bf-112">you can then run the code and try it out.</span></span>

<span data-ttu-id="0e4bf-113">Além disso, você pode fazer o download ou visualizar o projeto de exemplo completo no [Padrão de envio em lotes de funções personalizadas](https://github.com/OfficeDev/PnP-OfficeAddins/tree/master/Excel-custom-functions/Batching).</span><span class="sxs-lookup"><span data-stu-id="0e4bf-113">Also you can download or view the complete sample project at [Custom function batching pattern](https://github.com/OfficeDev/PnP-OfficeAddins/tree/master/Excel-custom-functions/Batching).</span></span> <span data-ttu-id="0e4bf-114">Se você quiser ver o código inteiro antes de ler mais, dê uma olhada no [arquivo de script](https://github.com/OfficeDev/PnP-OfficeAddins/blob/master/Excel-custom-functions/Batching/src/functions/functions.ts).</span><span class="sxs-lookup"><span data-stu-id="0e4bf-114">If you want to view the code in whole before reading any further, take a look at the [script file](https://github.com/OfficeDev/PnP-OfficeAddins/blob/master/Excel-custom-functions/Batching/src/functions/functions.ts).</span></span>

## <a name="create-the-batching-pattern-in-this-article"></a><span data-ttu-id="0e4bf-115">Crie o padrão de envio em lotes deste artigo</span><span class="sxs-lookup"><span data-stu-id="0e4bf-115">Create the batching pattern in this article</span></span>

<span data-ttu-id="0e4bf-116">Para configurar o envio em lotes para suas funções personalizadas, você precisará escrever três seções principais de código.</span><span class="sxs-lookup"><span data-stu-id="0e4bf-116">To set up batching for your custom functions you'll need to write three main sections of code.</span></span>

1. <span data-ttu-id="0e4bf-117">Uma operação push para adicionar uma nova operação ao lote de chamadas sempre que o Excel chamar sua função personalizada.</span><span class="sxs-lookup"><span data-stu-id="0e4bf-117">A push operation to add a new operation to the batch of calls each time Excel calls your custom function.</span></span>
2. <span data-ttu-id="0e4bf-118">Uma função para fazer o pedido remoto quando o lote estiver pronto.</span><span class="sxs-lookup"><span data-stu-id="0e4bf-118">A function to make the remote request when the batch is ready.</span></span>
3. <span data-ttu-id="0e4bf-119">Código do servidor para responder à solicitação em lote, calcular todos os resultados da operação e retornar os valores.</span><span class="sxs-lookup"><span data-stu-id="0e4bf-119">Server code to respond to the batch request, calculate all of the operation results, and return the values.</span></span>

<span data-ttu-id="0e4bf-120">Nas seções a seguir, você verá como construir o código com um exemplo de cada vez.</span><span class="sxs-lookup"><span data-stu-id="0e4bf-120">In the following sections you will be shown how to construct the code one example at a time.</span></span> <span data-ttu-id="0e4bf-121">Você adicionará cada exemplo de código ao seu arquivo **functions.ts**.</span><span class="sxs-lookup"><span data-stu-id="0e4bf-121">You'll add each code example to your functions.ts file.</span></span> <span data-ttu-id="0e4bf-122">É recomendável que você crie um novo projeto de funções personalizadas usando o gerador do Yo Office.</span><span class="sxs-lookup"><span data-stu-id="0e4bf-122">It's recommended you create a brand new custom functions project using yo office.</span></span> <span data-ttu-id="0e4bf-123">Para criar um novo projeto, confira [Começar a desenvolver funções personalizadas do Excel](../quickstarts/excel-custom-functions-quickstart.md) e use TypeScript em vez de JavaScript.</span><span class="sxs-lookup"><span data-stu-id="0e4bf-123">To create a new project see [Get started developing Excel custom functions](../quickstarts/excel-custom-functions-quickstart.md) and use TypeScript instead of JavaScript.</span></span>

## <a name="batch-each-call-to-your-custom-function"></a><span data-ttu-id="0e4bf-124">Agrupe cada chamada à sua função personalizada</span><span class="sxs-lookup"><span data-stu-id="0e4bf-124">Batch each call to your custom function</span></span>

<span data-ttu-id="0e4bf-125">Suas funções personalizadas funcionam chamando um serviço remoto para executar a operação e calcular o resultado de que precisam.</span><span class="sxs-lookup"><span data-stu-id="0e4bf-125">Your custom functions work by calling a remote service to perform the operation and calculate the result they need.</span></span> <span data-ttu-id="0e4bf-126">Isso fornece uma maneira de armazenar cada operação solicitada em um lote.</span><span class="sxs-lookup"><span data-stu-id="0e4bf-126">This provides a way for them to store each requested operation into a batch.</span></span> <span data-ttu-id="0e4bf-127">Mais tarde, você verá como criar uma função `_pushOperation` para agrupar as operações.</span><span class="sxs-lookup"><span data-stu-id="0e4bf-127">Later you'll see how to create a `_pushOperation` function to batch the operations.</span></span> <span data-ttu-id="0e4bf-128">Primeiro, dê uma olhada no exemplo de código a seguir para ver como chamar `_pushOperation` de sua função personalizada.</span><span class="sxs-lookup"><span data-stu-id="0e4bf-128">First, take a look at the following code example to see how to call `_pushOperation` from your custom function.</span></span>

<span data-ttu-id="0e4bf-129">No código a seguir, a função personalizada executa a divisão, mas depende de um serviço remoto para fazer o cálculo real.</span><span class="sxs-lookup"><span data-stu-id="0e4bf-129">In the following code, the custom function performs division but relies on a remote service to do the actual calculation.</span></span> <span data-ttu-id="0e4bf-130">Ela chama `_pushOperation` para reunir em lote a operação a outras operações para o serviço remoto.</span><span class="sxs-lookup"><span data-stu-id="0e4bf-130">It calls `_pushOperation` to batch the operation along with other operations to the remote service.</span></span> <span data-ttu-id="0e4bf-131">Nomeia a operação **div2**.</span><span class="sxs-lookup"><span data-stu-id="0e4bf-131">It names the operation **div2**.</span></span> <span data-ttu-id="0e4bf-132">Você pode usar qualquer esquema de nomenclatura desejado para operações, desde que o serviço remoto também esteja usando o mesmo esquema (mais informações sobre o serviço remoto posteriormente).</span><span class="sxs-lookup"><span data-stu-id="0e4bf-132">You can use any naming scheme you want for operations as long as the remote service is also using the same scheme (more on the remote service later).</span></span> <span data-ttu-id="0e4bf-133">Além disso, os argumentos que o serviço remoto precisará para executar a operação são passados.</span><span class="sxs-lookup"><span data-stu-id="0e4bf-133">Also, the arguments the remote service will need to run the operation are passed.</span></span>

### <a name="add-the-div2-custom-function-to-functionsts"></a><span data-ttu-id="0e4bf-134">Adicione a função customizada div2 ao functions.ts</span><span class="sxs-lookup"><span data-stu-id="0e4bf-134">Add the div2 custom function to functions.ts</span></span>

```typescript
/**
 * @CustomFunction
 * Divides two numbers using batching
 * @param dividend The number being divided
 * @param divisor The number the dividend is divided by
 * @returns The result of dividing the two numbers
 */
function div2(dividend: number, divisor: number) {
  return _pushOperation(
    "div2",
    [dividend, divisor]
  );
}

CustomFunctions.associate("DIV2", div2);
```

<span data-ttu-id="0e4bf-135">Em seguida, você definirá a matriz de lotes que armazenará todas as operações a serem passadas em uma chamada de rede.</span><span class="sxs-lookup"><span data-stu-id="0e4bf-135">Next, you will define the batch array which will store all operations to be passed in one network call.</span></span> <span data-ttu-id="0e4bf-136">O código a seguir mostra como definir uma interface descrevendo cada entrada de lote na matriz.</span><span class="sxs-lookup"><span data-stu-id="0e4bf-136">The following code shows how to define an interface describing each batch entry in the array.</span></span> <span data-ttu-id="0e4bf-137">A interface define uma operação, que é um nome de cadeia de caracteres da operação a ser executada.</span><span class="sxs-lookup"><span data-stu-id="0e4bf-137">The interface defines an operation, which is a string name of which operation to run.</span></span> <span data-ttu-id="0e4bf-138">Por exemplo, se você tivesse duas funções personalizadas nomeadas `multiply` e `divide`, você poderia reutilizá-las como nomes de operações em suas entradas de lote.</span><span class="sxs-lookup"><span data-stu-id="0e4bf-138">For example, if you had two custom functions named `multiply` and `divide`, you could reuse those as the operation names in your batch entries.</span></span> <span data-ttu-id="0e4bf-139">`args` manterá os argumentos que foram passados para sua função personalizada do Excel.</span><span class="sxs-lookup"><span data-stu-id="0e4bf-139">`args` will hold the arguments that were passed to your custom function from Excel.</span></span> <span data-ttu-id="0e4bf-140">E, finalmente, `resolve` ou `reject` armazenarão uma promessa contendo as informações que o serviço remoto retorna.</span><span class="sxs-lookup"><span data-stu-id="0e4bf-140">And finally, `resolve` or `reject` will store a promise holding the information the remote service returns.</span></span>

```typescript
interface IBatchEntry {
  operation: string;
  args: any[];
  resolve: (data: any) => void;
  reject: (error: Error) => void;
}
```

<span data-ttu-id="0e4bf-141">Em seguida, crie a matriz de lotes que usa a interface anterior.</span><span class="sxs-lookup"><span data-stu-id="0e4bf-141">Next, create the batch array that uses the previous interface.</span></span> <span data-ttu-id="0e4bf-142">Para controlar se um lote está programado ou não, crie uma variável `_isBatchedRequestSchedule`.</span><span class="sxs-lookup"><span data-stu-id="0e4bf-142">To track if a batch is scheduled or not, create an \`_isBatchedRequestSchedule variable.</span></span> <span data-ttu-id="0e4bf-143">Isso será importante mais tarde para o cronograma de chamadas em lote para o serviço remoto.</span><span class="sxs-lookup"><span data-stu-id="0e4bf-143">This will be important later for timing batch calls to the remote service.</span></span>

```typescript
const _batch: IBatchEntry[] = [];
let _isBatchedRequestScheduled = false;
```

<span data-ttu-id="0e4bf-144">Finalmente, quando o Excel chama sua função personalizada, você precisa enviar a operação para a matriz de lotes.</span><span class="sxs-lookup"><span data-stu-id="0e4bf-144">Finally when Excel calls your custom function, you need to push the operation into the batch array.</span></span> <span data-ttu-id="0e4bf-145">O código a seguir mostra como adicionar uma nova operação de uma função personalizada.</span><span class="sxs-lookup"><span data-stu-id="0e4bf-145">The following code shows how to add a new operation from a custom function.</span></span> <span data-ttu-id="0e4bf-146">Ele cria uma nova entrada de lote, cria uma nova promessa para resolver ou rejeitar a operação e envia a entrada para a matriz de lotes.</span><span class="sxs-lookup"><span data-stu-id="0e4bf-146">It creates a new batch entry, creates a new promise to resolve or reject the operation, and pushes the entry into the batch array.</span></span>

<span data-ttu-id="0e4bf-147">Esse código também verifica se um lote está programado.</span><span class="sxs-lookup"><span data-stu-id="0e4bf-147">This code also checks to see if a batch is scheduled.</span></span> <span data-ttu-id="0e4bf-148">Neste exemplo, cada lote está programado para ser executado a cada 100 ms.</span><span class="sxs-lookup"><span data-stu-id="0e4bf-148">In this example, each batch is scheduled to run every 100ms.</span></span> <span data-ttu-id="0e4bf-149">Você pode ajustar esse valor conforme necessário.</span><span class="sxs-lookup"><span data-stu-id="0e4bf-149">You can adjust this value as needed.</span></span> <span data-ttu-id="0e4bf-150">Valores mais altos resultam em lotes maiores sendo enviados ao serviço remoto e um tempo de espera maior para o usuário ver os resultados.</span><span class="sxs-lookup"><span data-stu-id="0e4bf-150">Higher values result in bigger batches being sent to the remote service, and a longer wait time for the user to see results.</span></span> <span data-ttu-id="0e4bf-151">Valores mais baixos tendem a enviar mais lotes para o serviço remoto, mas com um tempo de resposta rápido para os usuários.</span><span class="sxs-lookup"><span data-stu-id="0e4bf-151">Lower values tend to send more batches to the remote service, but with a quick response time for users.</span></span>

### <a name="add-the-pushoperation-function-to-functionsts"></a><span data-ttu-id="0e4bf-152">Adicione a função `_pushOperation` ao functions.ts</span><span class="sxs-lookup"><span data-stu-id="0e4bf-152">Add the `_pushOperation` function to functions.ts</span></span>

```typescript
function _pushOperation(op: string, args: any[]) {
  // Create an entry for your custom function.
  const invocationEntry: IBatchEntry = {
    operation: op, // e.g. sum
    args: args,
    resolve: undefined,
    reject: undefined,
  };

  // Create a unique promise for this invocation,
  // and save its resolve and reject functions into the invocation entry.
  const promise = new Promise((resolve, reject) => {
    invocationEntry.resolve = resolve;
    invocationEntry.reject = reject;
  });

  // Push the invocation entry into the next batch.
  _batch.push(invocationEntry);

  // If a remote request hasn't been scheduled yet,
  // schedule it after a certain timeout, e.g. 100 ms.
  if (!_isBatchedRequestScheduled) {
    _isBatchedRequestScheduled = true;
    setTimeout(_makeRemoteRequest, 100);
  }

  // Return the promise for this invocation.
  return promise;
}
```

## <a name="make-the-remote-request"></a><span data-ttu-id="0e4bf-153">Faça o pedido remoto</span><span class="sxs-lookup"><span data-stu-id="0e4bf-153">Make the remote request</span></span>

<span data-ttu-id="0e4bf-154">O objetivo da função `_makeRemoteRequest` é passar o lote de operações para o serviço remoto e, em seguida, retornar os resultados para cada função personalizada.</span><span class="sxs-lookup"><span data-stu-id="0e4bf-154">The purpose of the `_makeRemoteRequest` function is to pass the batch of operations to the remote service, and then return the results to each custom function.</span></span> <span data-ttu-id="0e4bf-155">Primeiro, ela cria uma cópia da matriz de lotes.</span><span class="sxs-lookup"><span data-stu-id="0e4bf-155">It first creates a copy of the batch array.</span></span> <span data-ttu-id="0e4bf-156">Isso permite que chamadas de função personalizadas simultâneas do Excel iniciem imediatamente o envio em lote em uma nova matriz.</span><span class="sxs-lookup"><span data-stu-id="0e4bf-156">This allows concurrent custom function calls from Excel to immediately begin batching in a new array.</span></span> <span data-ttu-id="0e4bf-157">A cópia é então transformada em uma matriz mais simples que não contém as informações de promessa.</span><span class="sxs-lookup"><span data-stu-id="0e4bf-157">The copy is then turned into a simpler array that does not contain the promise information.</span></span> <span data-ttu-id="0e4bf-158">Não faria sentido passar as promessas para um serviço remoto, uma vez que não funcionariam.</span><span class="sxs-lookup"><span data-stu-id="0e4bf-158">It wouldn't make sense to pass the promises to a remote service since they would not work.</span></span> <span data-ttu-id="0e4bf-159">`_makeRemoteRequest` irá rejeitar ou resolver cada promessa com base no que o serviço remoto retornar.</span><span class="sxs-lookup"><span data-stu-id="0e4bf-159">The \`_makeRemoteRequest will either reject or resolve each promise based on what the remote service returns.</span></span>

### <a name="add-the-following-makeremoterequest-method-to-functionsts"></a><span data-ttu-id="0e4bf-160">Adicione o seguinte método `_makeRemoteRequest` ao functions.ts</span><span class="sxs-lookup"><span data-stu-id="0e4bf-160">Add the following `_makeRemoteRequest` method to functions.ts</span></span>

```typescript
function _makeRemoteRequest() {
  // Copy the shared batch and allow the building of a new batch while you are waiting for a response.
  // Note the use of "splice" rather than "slice", which will modify the original _batch array
  // to empty it out.
  const batchCopy = _batch.splice(0, _batch.length);
  _isBatchedRequestScheduled = false;

  // Build a simpler request batch that only contains the arguments for each invocation.
  const requestBatch = batchCopy.map((item) => {
    return { operation: item.operation, args: item.args };
  });

  // Make the remote request.
  _fetchFromRemoteService(requestBatch)
    .then((responseBatch) => {
      // Match each value from the response batch to its corresponding invocation entry from the request batch,
      // and resolve the invocation promise with its corresponding response value.
      responseBatch.forEach((response, index) => {
        if (response.error) {
          batchCopy[index].reject(new Error(response.error));
        } else {
          console.log(response);
          batchCopy[index].resolve(response.result);
        }
      });
    });
}
```

### <a name="modify-makeremoterequest-for-your-own-solution"></a><span data-ttu-id="0e4bf-161">Modifique `_makeRemoteRequest` para sua própria solução</span><span class="sxs-lookup"><span data-stu-id="0e4bf-161">Modify `_makeRemoteRequest` for your own solution</span></span>

<span data-ttu-id="0e4bf-162">A função `_makeRemoteRequest` chama `_fetchFromRemoteService`, que, como você verá mais adiante, é apenas uma simulação representando o serviço remoto.</span><span class="sxs-lookup"><span data-stu-id="0e4bf-162">The `_makeRemoteRequest` function calls `_fetchFromRemoteService` which, as you'll see later, is just a mock representing the remote service.</span></span> <span data-ttu-id="0e4bf-163">Isso facilita estudar e executar o código neste artigo.</span><span class="sxs-lookup"><span data-stu-id="0e4bf-163">This makes it easier to study and run the code in this article.</span></span> <span data-ttu-id="0e4bf-164">Mas quando você quiser usar esse código para um serviço remoto real, faça as seguintes alterações:</span><span class="sxs-lookup"><span data-stu-id="0e4bf-164">But when you want to use this code for an actual remote service you should make the following changes:</span></span>

- <span data-ttu-id="0e4bf-165">Decida como serializar as operações em lote pela rede.</span><span class="sxs-lookup"><span data-stu-id="0e4bf-165">Decide how to serialize the batch operations over the network.</span></span> <span data-ttu-id="0e4bf-166">Por exemplo, você pode querer colocar a matriz em um corpo JSON.</span><span class="sxs-lookup"><span data-stu-id="0e4bf-166">For example, you may want to put the array into a JSON body.</span></span>
- <span data-ttu-id="0e4bf-167">Em vez de chamar `_fetchFromRemoteService`, você precisa fazer a chamada de rede real para o serviço remoto passando o lote de operações.</span><span class="sxs-lookup"><span data-stu-id="0e4bf-167">Instead of calling `_fetchFromRemoteService` you'll need to make the actual network call to the remote service passing the batch of operations.</span></span>

## <a name="process-the-batch-call-on-the-remote-service"></a><span data-ttu-id="0e4bf-168">Processar a chamada em lote no serviço remoto</span><span class="sxs-lookup"><span data-stu-id="0e4bf-168">Process the batch call on the remote service</span></span>

<span data-ttu-id="0e4bf-169">A última etapa é manipular a chamada em lote no serviço remoto.</span><span class="sxs-lookup"><span data-stu-id="0e4bf-169">The last step is to handle the batch call in the remote service.</span></span> <span data-ttu-id="0e4bf-170">O exemplo de código a seguir mostra a função `_fetchFromRemoteService`.</span><span class="sxs-lookup"><span data-stu-id="0e4bf-170">The following code sample shows the `_fetchFromRemoteService` function.</span></span> <span data-ttu-id="0e4bf-171">Essa função descompacta cada operação, executa a operação especificada e retorna os resultados.</span><span class="sxs-lookup"><span data-stu-id="0e4bf-171">This function unpacks each operation, performs the specified operation, and returns the results.</span></span> <span data-ttu-id="0e4bf-172">Para fins de aprendizado neste artigo, a função `_fetchFromRemoteService` foi projetada para ser executada em seu suplemento da Web e simular um serviço remoto.</span><span class="sxs-lookup"><span data-stu-id="0e4bf-172">For learning purposes in this article, the `_fetchFromRemoteService` function is designed to run in your web add-in and mock a remote service.</span></span> <span data-ttu-id="0e4bf-173">Você pode adicionar este código ao seu arquivo **functions.ts** para poder estudar e executar todo o código deste artigo sem precisar configurar um serviço remoto real.</span><span class="sxs-lookup"><span data-stu-id="0e4bf-173">You can add this code to your functions.ts file so that you can study and run all the code in this article without having to set up an actual remote service.</span></span>

### <a name="add-the-following-fetchfromremoteservice-function-to-functionsts"></a><span data-ttu-id="0e4bf-174">Adicione a seguinte função `_fetchFromRemoteService` ao functions.ts</span><span class="sxs-lookup"><span data-stu-id="0e4bf-174">Add the following `_fetchFromRemoteService` function to functions.ts</span></span>

```typescript
async function _fetchFromRemoteService(
  requestBatch: Array<{ operation: string, args: any[] }>
): Promise<IServerResponse[]> {
  // Simulate a slow network request to the server;
  await pause(1000);

  return requestBatch.map((request): IServerResponse => {
    const { operation, args } = request;

    try {
      if (operation === "div2") {
        // Divide the first argument by the second argument.
        return {
          result: args[0] / args[1]
        };
      } else if (operation === "mul2") {
        // Multiply the arguments for the given entry.
        const myresult = args[0] * args[1];
        console.log(myresult);
        return {
          result: myresult
        };
      } else {
        return {
          error: `Operation not supported: ${operation}`
        };
      }
    } catch (error) {
      return {
        error: `Operation failed: ${operation}`
      };
    }
  });
}

function pause(ms: number) {
  return new Promise((resolve) => setTimeout(resolve, ms));
}
```

### <a name="modify-fetchfromremoteservice-for-your-live-remote-service"></a><span data-ttu-id="0e4bf-175">Modifique `_fetchFromRemoteService` para o seu serviço remoto ao vivo</span><span class="sxs-lookup"><span data-stu-id="0e4bf-175">Modify `_fetchFromRemoteService` for your live remote service</span></span>

<span data-ttu-id="0e4bf-176">Para modificar a função `_fetchFromRemoteService` para que esta possa ser executada em seu serviço remoto ao vivo, faça as seguintes alterações:</span><span class="sxs-lookup"><span data-stu-id="0e4bf-176">To modify the `_fetchFromRemoteService` function to run in your live remote service, make the following changes:</span></span>

- <span data-ttu-id="0e4bf-177">Dependendo da plataforma do servidor (Node.js ou outros), mapeie a chamada de rede do cliente para essa função.</span><span class="sxs-lookup"><span data-stu-id="0e4bf-177">Depending on your server platform (Node.js or others) map the client network call to this function.</span></span>
- <span data-ttu-id="0e4bf-178">Remova a função `pause` que simula a latência da rede como parte da simulação.</span><span class="sxs-lookup"><span data-stu-id="0e4bf-178">Remove the `pause` function which simulates network latency as part of the mock.</span></span>
- <span data-ttu-id="0e4bf-179">Modifique a declaração da função para trabalhar com o parâmetro transmitido se o parâmetro for alterado para fins de rede.</span><span class="sxs-lookup"><span data-stu-id="0e4bf-179">Modify the function declaration to work with the parameter passed if the parameter is changed for network purposes.</span></span> <span data-ttu-id="0e4bf-180">Por exemplo, em vez de uma matriz, pode ser um corpo JSON de operações em lote a serem processadas.</span><span class="sxs-lookup"><span data-stu-id="0e4bf-180">For example, instead of an array, it may be a JSON body of batched operations to process.</span></span>
- <span data-ttu-id="0e4bf-181">Modifique a função para executar as operações (ou chame as funções que executam as operações).</span><span class="sxs-lookup"><span data-stu-id="0e4bf-181">Modify the function to perform the operations (or call functions that do the operations).</span></span>
- <span data-ttu-id="0e4bf-182">Aplique um mecanismo de autenticação apropriado.</span><span class="sxs-lookup"><span data-stu-id="0e4bf-182">Apply an appropriate authentication mechanism.</span></span> <span data-ttu-id="0e4bf-183">Certifique-se de que apenas os autores de chamada corretos possam acessar a função.</span><span class="sxs-lookup"><span data-stu-id="0e4bf-183">Ensure that only the correct callers can access the function.</span></span>
- <span data-ttu-id="0e4bf-184">Coloque o código no serviço remoto.</span><span class="sxs-lookup"><span data-stu-id="0e4bf-184">Place the code in the remote service.</span></span>

## <a name="next-steps"></a><span data-ttu-id="0e4bf-185">Próximas etapas</span><span class="sxs-lookup"><span data-stu-id="0e4bf-185">Next steps</span></span>
<span data-ttu-id="0e4bf-186">Saiba mais sobre [os vários parâmetros](custom-functions-parameter-options.md) que você pode usar nas suas funções personalizadas.</span><span class="sxs-lookup"><span data-stu-id="0e4bf-186">Learn about [the various parameters](custom-functions-parameter-options.md) you can use in your custom functions.</span></span> <span data-ttu-id="0e4bf-187">Ou, reveja as noções básicas sobre como fazer [uma chamada na Web através de um função personalizada](custom-functions-web-reqs.md).</span><span class="sxs-lookup"><span data-stu-id="0e4bf-187">Or review the basics behind making [a web call through a custom function](custom-functions-web-reqs.md).</span></span>

## <a name="see-also"></a><span data-ttu-id="0e4bf-188">Confira também</span><span class="sxs-lookup"><span data-stu-id="0e4bf-188">See also</span></span>

* [<span data-ttu-id="0e4bf-189">Valores voláteis nas funções</span><span class="sxs-lookup"><span data-stu-id="0e4bf-189">Volatile values in functions</span></span>](custom-functions-volatile.md)
* [<span data-ttu-id="0e4bf-190">Práticas recomendadas de funções personalizadas</span><span class="sxs-lookup"><span data-stu-id="0e4bf-190">Custom functions best practices</span></span>](custom-functions-best-practices.md)
* [<span data-ttu-id="0e4bf-191">Criar funções personalizadas no Excel</span><span class="sxs-lookup"><span data-stu-id="0e4bf-191">Create custom functions in Excel</span></span>](custom-functions-overview.md)
* [<span data-ttu-id="0e4bf-192">Tutorial de funções personalizadas do Excel</span><span class="sxs-lookup"><span data-stu-id="0e4bf-192">Excel custom functions tutorial</span></span>](../tutorials/excel-tutorial-create-custom-functions.md)
