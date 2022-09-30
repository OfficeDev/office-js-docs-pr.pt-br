---
ms.date: 09/09/2022
description: Reúna as funções personalizadas em lotes para reduzir as chamadas de rede para um serviço remoto.
title: Enviando em lote chamadas de função personalizada para um serviço remoto
ms.localizationpriority: medium
ms.openlocfilehash: f779351789350bbc591b1b5d7a975ff9f70cda26
ms.sourcegitcommit: cff5d3450f0c02814c1436f94cd1fc1537094051
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 09/30/2022
ms.locfileid: "68234918"
---
# <a name="batch-custom-function-calls-for-a-remote-service"></a>Chamadas de função personalizadas em lote para um serviço remoto

Se as suas funções personalizadas chamarem um serviço remoto, você poderá usar um padrão de envio em lotes para reduzir o número de chamadas de rede para o serviço remoto. Para reduzir a idas e voltas na rede, você reúne todas as chamadas em uma única chamada para o serviço da Web. Isso é ideal quando a planilha é recalculada.

Por exemplo, se alguém usou sua função personalizada em 100 células em uma planilha e depois recalculou a planilha, sua função personalizada seria executada 100 vezes e faria 100 chamadas de rede. Usando um padrão de envio em lotes, as chamadas podem ser combinadas para fazer todos os 100 cálculos em uma única chamada de rede.

[!include[Excel custom functions note](../includes/excel-custom-functions-note.md)]

## <a name="view-the-completed-sample"></a>Ver o exemplo concluído

Para exibir o exemplo concluído, siga este artigo e cole os exemplos de código em seu próprio projeto. Por exemplo, para criar um novo projeto de função personalizada para TypeScript, use o gerador [Yeoman para suplementos do Office](../develop/yeoman-generator-overview.md) e adicione todo o código deste artigo ao projeto. Execute o código e experimente-o.

Como alternativa, baixe ou exiba o projeto de exemplo completo no [padrão de envio em lote de funções personalizadas](https://github.com/OfficeDev/Office-Add-in-samples/tree/main/Excel-custom-functions/Batching). Se você quiser ver o código inteiro antes de ler mais, dê uma olhada no [arquivo de script](https://github.com/OfficeDev/Office-Add-in-samples/blob/main/Excel-custom-functions/Batching/src/functions/functions.js).

## <a name="create-the-batching-pattern-in-this-article"></a>Crie o padrão de envio em lotes deste artigo

Para configurar o envio em lotes para suas funções personalizadas, você precisará escrever três seções principais de código.

1. Uma [operação de push](#add-the-_pushoperation-function) para adicionar uma nova operação ao lote de chamadas sempre que o Excel chamar sua função personalizada.
2. Uma [função para fazer a solicitação remota](#make-the-remote-request) quando o lote estiver pronto.
3. [Código do servidor para responder à solicitação em](#process-the-batch-call-on-the-remote-service) lote, calcular todos os resultados da operação e retornar os valores.

Nas seções a seguir, você aprenderá a construir o código um exemplo de cada vez. É recomendável que você crie um novo projeto de funções personalizadas usando o gerador [Yeoman para o gerador de Suplementos do Office](../develop/yeoman-generator-overview.md) . Para criar um novo projeto, consulte [Introdução ao desenvolvimento de funções personalizadas do Excel](../quickstarts/excel-custom-functions-quickstart.md). Você pode usar TypeScript ou JavaScript.

## <a name="batch-each-call-to-your-custom-function"></a>Agrupe cada chamada à sua função personalizada

Suas funções personalizadas funcionam chamando um serviço remoto para executar a operação e calcular o resultado de que precisam. Isso fornece uma maneira de armazenar cada operação solicitada em um lote. Mais tarde, você verá como criar uma função `_pushOperation` para agrupar as operações. Primeiro, dê uma olhada no exemplo de código a seguir para ver como chamar `_pushOperation` de sua função personalizada.

No código a seguir, a função personalizada executa a divisão, mas depende de um serviço remoto para fazer o cálculo real. Ela chama `_pushOperation` para reunir em lote a operação a outras operações para o serviço remoto. Nomeia a operação **div2**. Você pode usar qualquer esquema de nomenclatura desejado para operações, desde que o serviço remoto também esteja usando o mesmo esquema (mais informações sobre o serviço remoto posteriormente). Além disso, os argumentos que o serviço remoto precisará para executar a operação são passados.

### <a name="add-the-div2-custom-function"></a>Adicionar a função personalizada div2

Adicione o código a seguir ao **arquivofunctions.js** **ou functions.ts** (dependendo se você usou JavaScript ou TypeScript).

```javascript
/**
 * Divides two numbers using batching
 * @CustomFunction
 * @param dividend The number being divided
 * @param divisor The number the dividend is divided by
 * @returns The result of dividing the two numbers
 */
function div2(dividend, divisor) {
  return _pushOperation("div2", [dividend, divisor]);
}
```

### <a name="add-global-variables-for-tracking-batch-requests"></a>Adicionar variáveis globais para acompanhar solicitações em lote

Em seguida, adicione duas variáveis globais **ao arquivofunctions.js** **ou functions.ts** . `_isBatchedRequestScheduled` é importante posteriormente para cronometrização de chamadas em lote para o serviço remoto.

```javascript
let _batch = [];
let _isBatchedRequestScheduled = false;
```

### <a name="add-the-_pushoperation-function"></a>Adicionar a `_pushOperation` função

Quando o Excel chama sua função personalizada, você precisa enviar a operação por push para a matriz de lote. O código **_pushOperation** função a seguir mostra como adicionar uma nova operação de uma função personalizada. Ele cria uma nova entrada de lote, cria uma nova promessa para resolver ou rejeitar a operação e envia a entrada para a matriz de lotes.

Esse código também verifica se um lote está programado. Neste exemplo, cada lote está programado para ser executado a cada 100 ms. Você pode ajustar esse valor conforme necessário. Valores mais altos resultam em lotes maiores sendo enviados ao serviço remoto e um tempo de espera maior para o usuário ver os resultados. Valores mais baixos tendem a enviar mais lotes para o serviço remoto, mas com um tempo de resposta rápido para os usuários.

A função cria um **objeto invocationEntry** que contém o nome da cadeia de caracteres da operação a ser executada. Por exemplo, se você tivesse duas funções personalizadas nomeadas `multiply` e `divide`, você poderia reutilizá-las como nomes de operações em suas entradas de lote. `args` contém os argumentos que foram passados para sua função personalizada do Excel. E, por fim, `resolve` ou `reject` os métodos armazenam uma promessa que contém as informações retornadas pelo serviço remoto.

Adicione o código a seguir ao **arquivofunctions.js** **ou functions.ts** .

```javascript
// This function encloses your custom functions as individual entries,
// which have some additional properties so you can keep track of whether or not
// a request has been resolved or rejected.
function _pushOperation(op, args) {
  // Create an entry for your custom function.
  console.log("pushOperation");
  const invocationEntry = {
    operation: op, // e.g., sum
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
  // schedule it after a certain timeout, e.g., 100 ms.
  if (!_isBatchedRequestScheduled) {
    console.log("schedule remote request");
    _isBatchedRequestScheduled = true;
    setTimeout(_makeRemoteRequest, 100);
  }

  // Return the promise for this invocation.
  return promise;
}
```

## <a name="make-the-remote-request"></a>Faça o pedido remoto

O objetivo da função `_makeRemoteRequest` é passar o lote de operações para o serviço remoto e, em seguida, retornar os resultados para cada função personalizada. Primeiro, ela cria uma cópia da matriz de lotes. Isso permite que chamadas de função personalizadas simultâneas do Excel iniciem imediatamente o envio em lote em uma nova matriz. A cópia é então transformada em uma matriz mais simples que não contém as informações de promessa. Não faria sentido passar as promessas para um serviço remoto, uma vez que não funcionariam. `_makeRemoteRequest` irá rejeitar ou resolver cada promessa com base no que o serviço remoto retornar.

Adicione o código a seguir ao **arquivofunctions.js** **ou functions.ts** .

```javascript
// This is a private helper function, used only within your custom function add-in.
// You wouldn't call _makeRemoteRequest in Excel, for example.
// This function makes a request for remote processing of the whole batch,
// and matches the response batch to the request batch.
function _makeRemoteRequest() {
  // Copy the shared batch and allow the building of a new batch while you are waiting for a response.
  // Note the use of "splice" rather than "slice", which will modify the original _batch array
  // to empty it out.
  try{
  console.log("makeRemoteRequest");
  const batchCopy = _batch.splice(0, _batch.length);
  _isBatchedRequestScheduled = false;

  // Build a simpler request batch that only contains the arguments for each invocation.
  const requestBatch = batchCopy.map((item) => {
    return { operation: item.operation, args: item.args };
  });
  console.log("makeRemoteRequest2");
  // Make the remote request.
  _fetchFromRemoteService(requestBatch)
    .then((responseBatch) => {
      console.log("responseBatch in fetchFromRemoteService");
      // Match each value from the response batch to its corresponding invocation entry from the request batch,
      // and resolve the invocation promise with its corresponding response value.
      responseBatch.forEach((response, index) => {
        if (response.error) {
          batchCopy[index].reject(new Error(response.error));
          console.log("rejecting promise");
        } else {
          console.log("fulfilling promise");
          console.log(response);

          batchCopy[index].resolve(response.result);
        }
      });
    });
    console.log("makeRemoteRequest3");
  } catch (error) {
    console.log("error name:" + error.name);
    console.log("error message:" + error.message);
    console.log(error);
  }
}
```

### <a name="modify-_makeremoterequest-for-your-own-solution"></a>Modifique `_makeRemoteRequest` para sua própria solução

A função `_makeRemoteRequest` chama `_fetchFromRemoteService`, que, como você verá mais adiante, é apenas uma simulação representando o serviço remoto. Isso facilita estudar e executar o código neste artigo. Mas quando você quiser usar esse código para um serviço remoto real, deverá fazer as alterações a seguir.

- Decida como serializar as operações em lote pela rede. Por exemplo, você pode querer colocar a matriz em um corpo JSON.
- Em vez de chamar `_fetchFromRemoteService`, você precisa fazer a chamada de rede real para o serviço remoto passando o lote de operações.

## <a name="process-the-batch-call-on-the-remote-service"></a>Processar a chamada em lote no serviço remoto

A última etapa é manipular a chamada em lote no serviço remoto. O exemplo de código a seguir mostra a função `_fetchFromRemoteService`. Essa função descompacta cada operação, executa a operação especificada e retorna os resultados. Para fins de aprendizado neste artigo, a função `_fetchFromRemoteService` foi projetada para ser executada em seu suplemento da Web e simular um serviço remoto. Você pode adicionar esse código ao arquivo **functions.js** **ou functions.ts** para que possa estudar e executar todo o código neste artigo sem precisar configurar um serviço remoto real.

Adicione o código a seguir ao **arquivofunctions.js** **ou functions.ts** .

```javascript
// This function simulates the work of a remote service. Because each service
// differs, you will need to modify this function appropriately to work with the service you are using. 
// This function takes a batch of argument sets and returns a promise that may contain a batch of values.
// NOTE: When implementing this function on a server, also apply an appropriate authentication mechanism
//       to ensure only the correct callers can access it.
async function _fetchFromRemoteService(requestBatch) {
  // Simulate a slow network request to the server.
  console.log("_fetchFromRemoteService");
  await pause(1000);
  console.log("postpause");
  return requestBatch.map((request) => {
    console.log("requestBatch server side");
    const { operation, args } = request;

    try {
      if (operation === "div2") {
        // Divide the first argument by the second argument.
        return {
          result: args[0] / args[1]
        };
      } else if (operation === "mul2") {
        // Multiply the arguments for the given entry.
        const myResult = args[0] * args[1];
        console.log(myResult);
        return {
          result: myResult
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

function pause(ms) {
  console.log("pause");
  return new Promise((resolve) => setTimeout(resolve, ms));
}
```

### <a name="modify-_fetchfromremoteservice-for-your-live-remote-service"></a>Modifique `_fetchFromRemoteService` para o seu serviço remoto ao vivo

Para modificar a `_fetchFromRemoteService` função a ser executada em seu serviço remoto dinâmico, faça as seguintes alterações.

- Dependendo da plataforma do servidor (Node.js ou outros), mapeie a chamada de rede do cliente para essa função.
- Remova a função `pause` que simula a latência da rede como parte da simulação.
- Modifique a declaração da função para trabalhar com o parâmetro transmitido se o parâmetro for alterado para fins de rede. Por exemplo, em vez de uma matriz, pode ser um corpo JSON de operações em lote a serem processadas.
- Modifique a função para executar as operações (ou chame as funções que executam as operações).
- Aplique um mecanismo de autenticação apropriado. Certifique-se de que apenas os autores de chamada corretos possam acessar a função.
- Coloque o código no serviço remoto.

## <a name="next-steps"></a>Próximas etapas

Saiba mais sobre [os vários parâmetros](custom-functions-parameter-options.md) que você pode usar nas suas funções personalizadas. Ou, reveja as noções básicas sobre como fazer [uma chamada na Web através de um função personalizada](custom-functions-web-reqs.md).

## <a name="see-also"></a>Confira também

- [Valores voláteis nas funções](custom-functions-volatile.md)
- [Criar funções personalizadas no Excel](custom-functions-overview.md)
- [Tutorial de funções personalizadas do Excel](../tutorials/excel-tutorial-create-custom-functions.md)
