---
ms.date: 05/02/2022
description: Solicitar, transmitir e cancelar o streaming de dados externos para sua pasta de trabalho com funções personalizadas no Excel.
title: Receber e tratar dados com funções personalizadas
ms.localizationpriority: medium
ms.openlocfilehash: fbe319e79d4cded5fe4b37ce5a654e633996f22a
ms.sourcegitcommit: b6a3815a1ad17f3522ca35247a3fd5d7105e174e
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 07/22/2022
ms.locfileid: "66958542"
---
# <a name="receive-and-handle-data-with-custom-functions"></a>Receber e tratar dados com funções personalizadas

Uma das maneiras pelas quais as funções personalizadas aprimoram o poder do Excel é recebendo dados de locais diferentes da pasta de trabalho, como a Web ou um servidor (por meio de [WebSockets](https://developer.mozilla.org/docs/Web/API/WebSockets_API)). É possível solicitar dados externos através de uma API como [ `Fetch` ](https://developer.mozilla.org/docs/Web/API/Fetch_API) ou usando `XmlHttpRequest` [(XHR)](https://developer.mozilla.org/docs/Web/API/XMLHttpRequest), uma API Web padrão que envia solicitações HTTP para interagir com os servidores.

[!include[Excel custom functions note](../includes/excel-custom-functions-note.md)]

![GIF de uma função personalizada que transmite o tempo de uma API.](../images/custom-functions-web-api.gif)

## <a name="functions-that-return-data-from-external-sources"></a>Funções que retornam os dados de fontes externas

Se uma função personalizada recupera dados de uma fonte externa como na web, ela deve:

1. Retornar um [JavaScript `Promise`](https://developer.mozilla.org/docs/Web/JavaScript/Reference/Global_Objects/Promise) para o Excel.
2. Resolva com `Promise` o valor final usando a função de retorno de chamada.

### <a name="fetch-example"></a>Exemplo de busca

No exemplo de código a seguir, `webRequest` a função alcança uma API externa hipotética que rastreia o número de pessoas atualmente na Estação Espacial Internacional. A função retorna um JavaScript e `Promise` usa para `fetch` solicitar informações da API hipotética. Os dados resultantes são transformados em JSON `names` e a propriedade é convertida em uma cadeia de caracteres, que é usada para resolver a promessa.

Ao desenvolver suas próprias funções, talvez você queira executar uma ação caso a solicitação da Web não tenha sido concluída de maneira oportuna ou considere [o envio de várias solicitações](custom-functions-batching.md)da API.

```JS
/**
 * Requests the names of the people currently on the International Space Station.
 * Note: This function requests data from a hypothetical URL. In practice, replace the URL with a data source for your scenario.
 * @customfunction
 */
function webRequest() {
  let url = "https://www.contoso.com/NumberOfPeopleInSpace"; // This is a hypothetical URL.
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

> [!NOTE]
> Usar `fetch` evita retornos de chamada aninhados e pode ser preferível do XHR em alguns casos.

### <a name="xhr-example"></a>Exemplo de XHR

No exemplo de código a seguir, `getStarCount` a função chama a API do Github para descobrir a quantidade de estrelas fornecidas ao repositório de um usuário específico. Essa é uma função assíncrona que retorna um JavaScript `Promise`. Quando os dados são obtidos da chamada à Web, a promessa é resolvida, o que retorna os dados para a célula.

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

## <a name="make-a-streaming-function"></a>Faça uma função de streaming

Funções personalizadas de streaming permitem a saída de dados para células que atualizam repetidamente, sem a necessidade de um usuário explicitamente atualizar coisa alguma. Isso pode ser útil para verificar dados ativos de um serviço online, como a função no [tutorial de funções personalizadas](../tutorials/excel-tutorial-create-custom-functions.md).

Para declarar uma função de streaming, você pode usar qualquer uma das duas opções a seguir.

- A `@streaming` marca.
- O `CustomFunctions.StreamingInvocation` parâmetro de invocação.

O exemplo a seguir é uma função personalizada que adiciona um número ao resultado a cada segundo. Observe o seguinte sobre este código.

- O Excel exibe cada valor novo automaticamente usando o método `setResult`.
- O segundo parâmetro de entrada, `invocation`, não é exibido para os usuários finais no Excel quando eles selecionam a função no menu de preenchimento automático.
- O `onCanceled` retorno de chamada define a função que é executada quando a função é cancelada.
- O streaming não está necessariamente vinculado a fazer uma solicitação da Web. Nesse caso, a função não está fazendo uma solicitação da Web, mas ainda está obtendo dados em intervalos definidos, portanto, ela requer o uso do parâmetro de streaming `invocation` .

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
```

## <a name="cancel-a-function"></a>Cancelar uma função

O Excel cancela a execução de uma função nas situações a seguir.

- Quando o usuário edita ou exclui uma célula que faz referência à função.
- Quando é alterado um dos argumentos (entradas) para a função. Nesse caso, uma nova chamada de função é disparada, seguindo o cancelamento.
- Quando o usuário aciona manualmente um recálculo. Nesse caso, uma nova chamada de função é disparada, seguindo o cancelamento.

Você também pode considerar a definição de um valor de streaming padrão para lidar com os casos em que uma solicitação for feita, mas você está offline.

> [!NOTE]
> Também há uma categoria de funções chamadas funções canceláveis, e elas não estão relacionadas _a funções_ de streaming. Somente funções personalizadas assíncronas que retornam um valor são canceláveis. Funções canceláveis permitem que uma solicitação da Web seja encerrada no meio de uma solicitação, usando um [`CancelableInvocation`](/javascript/api/custom-functions-runtime/customfunctions.cancelableinvocation) para decidir o que fazer após o cancelamento. Declare uma função cancelável usando a tag `@cancelable`.

### <a name="use-an-invocation-parameter"></a>Usar um parâmetro de invocação

O parâmetro `invocation` é o último parâmetro de qualquer função personalizada por padrão. O `invocation` parâmetro fornece contexto sobre a célula (como seu endereço e conteúdo) e permite que você use `setResult` `onCanceled` o método e o evento para definir o que uma função faz quando ela transmite (`setResult`) ou é cancelada (`onCanceled`).

Se você estiver usando TypeScript, o manipulador de invocação precisará ser do tipo [`CustomFunctions.StreamingInvocation`](/javascript/api/custom-functions-runtime/customfunctions.streaminginvocation) ou [`CancelableInvocation`](/javascript/api/custom-functions-runtime/customfunctions.cancelableinvocation).

## <a name="receiving-data-via-websockets"></a>Como receber dados por meio de WebSockets

Em uma função personalizada, é possível usar [WebSockets](https://developer.mozilla.org/docs/Web/API/WebSockets_API) para trocar dados por meio de uma conexão persistente com um servidor. Usando WebSockets, sua função personalizada pode abrir uma conexão com um servidor e receber automaticamente mensagens do servidor quando determinados eventos ocorrerem, sem precisar sondar explicitamente o servidor para obter dados.

### <a name="websockets-example"></a>Exemplo de WebSockets

O código de exemplo a seguir estabelece uma conexão WebSocket e registra cada mensagem de entrada do servidor.

```js
let ws = new WebSocket('wss://bundles.office.com');

ws.onmessage(message) {
    console.log(`Received: ${message}`);
}

ws.onerror(error){
    console.err(`Failed: ${error}`);
}
```

## <a name="next-steps"></a>Próximas etapas

- Saiba mais sobre [diferentes tipos de parâmetros que as suas funções podem usar](custom-functions-parameter-options.md).
- Descubra como [agrupar várias chamadas de API](custom-functions-batching.md).

## <a name="see-also"></a>Confira também

- [Valores voláteis nas funções](custom-functions-volatile.md)
- [Criar metadados JSON para funções personalizadas](custom-functions-json-autogeneration.md)
- [Criar manualmente metadados JSON para funções personalizadas](custom-functions-json.md)
- [Criar funções personalizadas no Excel](custom-functions-overview.md)
- [Tutorial de funções personalizadas do Excel](../tutorials/excel-tutorial-create-custom-functions.md)
