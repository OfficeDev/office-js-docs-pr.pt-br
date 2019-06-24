---
ms.date: 06/17/2019
description: Reúna as funções personalizadas em lotes para reduzir as chamadas de rede para um serviço remoto.
title: Enviando em lote chamadas de função personalizada para um serviço remoto
localization_priority: Priority
ms.openlocfilehash: aa1b9c956c0f54a4d59e49ca157dd67c8349b143
ms.sourcegitcommit: 382e2735a1295da914f2bfc38883e518070cec61
ms.translationtype: HT
ms.contentlocale: pt-BR
ms.lasthandoff: 06/21/2019
ms.locfileid: "35127937"
---
# <a name="batching-custom-function-calls-for-a-remote-service"></a>Enviando em lote chamadas de função personalizada para um serviço remoto

Se as suas funções personalizadas chamarem um serviço remoto, você poderá usar um padrão de envio em lotes para reduzir o número de chamadas de rede para o serviço remoto. Para reduzir a idas e voltas na rede, você reúne todas as chamadas em uma única chamada para o serviço da Web. Isso é ideal quando a planilha é recalculada.

Por exemplo, se alguém usou sua função personalizada em 100 células em uma planilha e depois recalculou a planilha, sua função personalizada seria executada 100 vezes e faria 100 chamadas de rede. Usando um padrão de envio em lotes, as chamadas podem ser combinadas para fazer todos os 100 cálculos em uma única chamada de rede.

[!include[Excel custom functions note](../includes/excel-custom-functions-note.md)]

## <a name="view-the-completed-sample"></a>Ver o exemplo concluído

Você pode seguir este artigo e colar os exemplos de código em seu próprio projeto. Por exemplo, você pode usar o [gerador do Yo Office](https://github.com/OfficeDev/generator-office) para criar um novo projeto de função personalizada para TypeScript e, em seguida, adicionar todo o código deste artigo ao projeto. Você pode então executar o código e experimentá-lo.

Além disso, você pode fazer o download ou visualizar o projeto de exemplo completo no [Padrão de envio em lotes de funções personalizadas](https://github.com/OfficeDev/PnP-OfficeAddins/tree/master/Excel-custom-functions/Batching). Se você quiser ver o código inteiro antes de ler mais, dê uma olhada no [arquivo de script](https://github.com/OfficeDev/PnP-OfficeAddins/blob/master/Excel-custom-functions/Batching/src/functions/functions.ts).

## <a name="create-the-batching-pattern-in-this-article"></a>Crie o padrão de envio em lotes deste artigo

Para configurar o envio em lotes para suas funções personalizadas, você precisará escrever três seções principais de código.

1. Uma operação push para adicionar uma nova operação ao lote de chamadas sempre que o Excel chamar sua função personalizada.
2. Uma função para fazer o pedido remoto quando o lote estiver pronto.
3. Código do servidor para responder à solicitação em lote, calcular todos os resultados da operação e retornar os valores.

Nas seções a seguir, você verá como construir o código com um exemplo de cada vez. Você adicionará cada exemplo de código ao seu arquivo **functions.ts**. É recomendável que você crie um novo projeto de funções personalizadas usando o gerador do Yo Office. Para criar um novo projeto, confira [Começar a desenvolver funções personalizadas do Excel](../quickstarts/excel-custom-functions-quickstart.md) e use TypeScript em vez de JavaScript.

## <a name="batch-each-call-to-your-custom-function"></a>Agrupe cada chamada à sua função personalizada

Suas funções personalizadas funcionam chamando um serviço remoto para executar a operação e calcular o resultado de que precisam. Isso fornece uma maneira de armazenar cada operação solicitada em um lote. Mais tarde, você verá como criar uma função `_pushOperation` para agrupar as operações. Primeiro, dê uma olhada no exemplo de código a seguir para ver como chamar `_pushOperation` de sua função personalizada.

No código a seguir, a função personalizada executa a divisão, mas depende de um serviço remoto para fazer o cálculo real. Ela chama `_pushOperation` para reunir em lote a operação a outras operações para o serviço remoto. Nomeia a operação **div2**. Você pode usar qualquer esquema de nomenclatura desejado para operações, desde que o serviço remoto também esteja usando o mesmo esquema (mais informações sobre o serviço remoto posteriormente). Além disso, os argumentos que o serviço remoto precisará para executar a operação são passados.

### <a name="add-the-div2-custom-function-to-functionsts"></a>Adicione a função customizada div2 ao functions.ts

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

Em seguida, você definirá a matriz de lotes que armazenará todas as operações a serem passadas em uma chamada de rede. O código a seguir mostra como definir uma interface descrevendo cada entrada de lote na matriz. A interface define uma operação, que é um nome de cadeia de caracteres da operação a ser executada. Por exemplo, se você tivesse duas funções personalizadas nomeadas `multiply` e `divide`, você poderia reutilizá-las como nomes de operações em suas entradas de lote. `args` manterá os argumentos que foram passados para sua função personalizada do Excel. E, finalmente, `resolve` ou `reject` armazenarão uma promessa contendo as informações que o serviço remoto retorna.

```typescript
interface IBatchEntry {
  operation: string;
  args: any[];
  resolve: (data: any) => void;
  reject: (error: Error) => void;
}
```

Em seguida, crie a matriz de lotes que usa a interface anterior. Para controlar se um lote está programado ou não, crie uma variável `_isBatchedRequestSchedule`. Isso será importante mais tarde para o cronograma de chamadas em lote para o serviço remoto.

```typescript
const _batch: IBatchEntry[] = [];
let _isBatchedRequestScheduled = false;
```

Finalmente, quando o Excel chama sua função personalizada, você precisa enviar a operação para a matriz de lotes. O código a seguir mostra como adicionar uma nova operação de uma função personalizada. Ele cria uma nova entrada de lote, cria uma nova promessa para resolver ou rejeitar a operação e envia a entrada para a matriz de lotes.

Esse código também verifica se um lote está programado. Neste exemplo, cada lote está programado para ser executado a cada 100 ms. Você pode ajustar esse valor conforme necessário. Valores mais altos resultam em lotes maiores sendo enviados ao serviço remoto e um tempo de espera maior para o usuário ver os resultados. Valores mais baixos tendem a enviar mais lotes para o serviço remoto, mas com um tempo de resposta rápido para os usuários.

### <a name="add-the-pushoperation-function-to-functionsts"></a>Adicione a função `_pushOperation` ao functions.ts

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

## <a name="make-the-remote-request"></a>Faça o pedido remoto

O objetivo da função `_makeRemoteRequest` é passar o lote de operações para o serviço remoto e, em seguida, retornar os resultados para cada função personalizada. Primeiro, ela cria uma cópia da matriz de lotes. Isso permite que chamadas de função personalizadas simultâneas do Excel iniciem imediatamente o envio em lote em uma nova matriz. A cópia é então transformada em uma matriz mais simples que não contém as informações de promessa. Não faria sentido passar as promessas para um serviço remoto, uma vez que não funcionariam. `_makeRemoteRequest` irá rejeitar ou resolver cada promessa com base no que o serviço remoto retornar.

### <a name="add-the-following-makeremoterequest-method-to-functionsts"></a>Adicione o seguinte método `_makeRemoteRequest` ao functions.ts

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

### <a name="modify-makeremoterequest-for-your-own-solution"></a>Modifique `_makeRemoteRequest` para sua própria solução

A função `_makeRemoteRequest` chama `_fetchFromRemoteService`, que, como você verá mais adiante, é apenas uma simulação representando o serviço remoto. Isso facilita estudar e executar o código neste artigo. Mas quando você quiser usar esse código para um serviço remoto real, faça as seguintes alterações:

- Decida como serializar as operações em lote pela rede. Por exemplo, você pode querer colocar a matriz em um corpo JSON.
- Em vez de chamar `_fetchFromRemoteService`, você precisa fazer a chamada de rede real para o serviço remoto passando o lote de operações.

## <a name="process-the-batch-call-on-the-remote-service"></a>Processar a chamada em lote no serviço remoto

A última etapa é manipular a chamada em lote no serviço remoto. O exemplo de código a seguir mostra a função `_fetchFromRemoteService`. Essa função descompacta cada operação, executa a operação especificada e retorna os resultados. Para fins de aprendizado neste artigo, a função `_fetchFromRemoteService` foi projetada para ser executada em seu suplemento da Web e simular um serviço remoto. Você pode adicionar este código ao seu arquivo **functions.ts** para poder estudar e executar todo o código deste artigo sem precisar configurar um serviço remoto real.

### <a name="add-the-following-fetchfromremoteservice-function-to-functionsts"></a>Adicione a seguinte função `_fetchFromRemoteService` ao functions.ts

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

### <a name="modify-fetchfromremoteservice-for-your-live-remote-service"></a>Modifique `_fetchFromRemoteService` para o seu serviço remoto ao vivo

Para modificar a função `_fetchFromRemoteService` para que esta possa ser executada em seu serviço remoto ao vivo, faça as seguintes alterações:

- Dependendo da plataforma do servidor (Node.js ou outros), mapeie a chamada de rede do cliente para essa função.
- Remova a função `pause` que simula a latência da rede como parte da simulação.
- Modifique a declaração da função para trabalhar com o parâmetro transmitido se o parâmetro for alterado para fins de rede. Por exemplo, em vez de uma matriz, pode ser um corpo JSON de operações em lote a serem processadas.
- Modifique a função para executar as operações (ou chame as funções que executam as operações).
- Aplique um mecanismo de autenticação apropriado. Certifique-se de que apenas os autores de chamada corretos possam acessar a função.
- Coloque o código no serviço remoto.

## <a name="next-steps"></a>Próximas etapas
Saiba mais sobre [os vários parâmetros](custom-functions-parameter-options.md) que você pode usar nas suas funções personalizadas. Ou, reveja as noções básicas sobre como fazer [uma chamada na Web através de um função personalizada](custom-functions-web-reqs.md).

## <a name="see-also"></a>Confira também

* [Valores voláteis nas funções](custom-functions-volatile.md)
* [Práticas recomendadas de funções personalizadas](custom-functions-best-practices.md)
* [Criar funções personalizadas no Excel](custom-functions-overview.md)
* [Tutorial de funções personalizadas do Excel](../tutorials/excel-tutorial-create-custom-functions.md)
