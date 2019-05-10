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
# <a name="receive-and-handle-data-with-custom-functions"></a>Receber e tratar dados com funções personalizadas

Uma das maneiras pelas quais as funções personalizadas aprimoram o poder do Excel é através do recebimento de dados de outros locais diferente da pasta de trabalho, como a Web ou um servidor (por meio de WebSockets). As funções personalizadas podem solicitar dados por meio de XHR e solicitações `fetch`, bem como transmitir esses dados em tempo real.

[!include[Excel custom functions note](../includes/excel-custom-functions-note.md)]

A documentação a seguir ilustra alguns exemplos de solicitações da web, mas para criar uma função de transmissão para você, experimente o [Tutorial de funções personalizadas](https://docs.microsoft.com/office/dev/add-ins/tutorials/excel-tutorial-create-custom-functions?tabs=excel-windows).

## <a name="functions-that-return-data-from-external-sources"></a>Funções que retornam os dados de fontes externas

Se uma função personalizada recupera dados de uma fonte externa como na web, ela deve:

1. Retornar uma Promise do JavaScript para o Excel.
2. Resolva a promessa com o uso da função retorno de chamada de valor final.

É possível solicitar dados externos através de uma API como [ `Fetch` ](https://developer.mozilla.org/pt-BR/docs/Web/API/Fetch_API) ou usando `XmlHttpRequest` [(XHR)](https://developer.mozilla.org/pt-BR/docs/Web/API/XMLHttpRequest), uma API Web padrão que envia solicitações HTTP para interagir com os servidores.

No tempo de execução das funções personalizadas, o XHR implementa medidas de segurança adicionais solicitando uma [Política de mesma origem](https://developer.mozilla.org/pt-BR/docs/Web/Security/Same-origin_policy) ou um simples [CORS](https://www.w3.org/TR/cors/).

Observe que uma implementação CORS simples não pode usar cookies e é compatível apenas com métodos simples (GET, HEAD, POST). A CORS simples aceita cabeçalhos simples com nomes de campos `Accept`, `Accept-Language`, `Content-Language`. Você também pode usar um cabeçalho de tipo de conteúdo no CORS simples, desde que o tipo de conteúdo seja `application/x-www-form-urlencoded`, `text/plain`, ou `multipart/form-data`.

### <a name="xhr-example"></a>Exemplo de XHR

No código de exemplo a seguir, a função **getTemperature** chama a função sendWebRequest  para obter a temperatura de uma área específica, de acordo com a ID do termômetro. A função sendWebRequest usa XHR para emitir uma solicitação GET para um ponto de extremidade que fornece os dados.

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

Para outro exemplo de solicitação XHR com mais contexto, confira a função`getFile` dentro [deste arquivo](https://github.com/OfficeDev/Office-Add-in-JavaScript-FileDownload/blob/master/FileDownloadSampleWeb/Home.js) no repositório Github [Office-Add-in-JavaScript-FileDownload](https://github.com/OfficeDev/Office-Add-in-JavaScript-FileDownload).

### <a name="fetch-example"></a>Exemplo de busca

No seguinte exemplo de código, a função `stockPriceStream` usa um símbolo de cotação da bolsa para acessar o preço de uma ação a cada 1000 milissegundos. Para saber mais sobre este exemplo, confira o [Tutorial de funções personalizadas](https://docs.microsoft.com/office/dev/add-ins/tutorials/excel-tutorial-create-custom-functions?tabs=excel-windows#create-a-streaming-asynchronous-custom-function).

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

## <a name="receive-data-via-websockets"></a>Receber dados por meio de WebSockets

Em uma função personalizada, é possível usar WebSockets para trocar dados por meio de uma conexão persistente com um servidor. Usando WebSockets, a função personalizada pode abrir uma conexão com um servidor e, em seguida, receber mensagens do servidor automaticamente, quando determinados eventos ocorrerem, sem precisar consultar explicitamente os dados do servidor.

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

## <a name="stream-and-cancel-functions"></a>Funções stream e cancel

Funções personalizadas de streaming permitem a saída de dados para células que atualizam repetidamente, sem a necessidade de um usuário explicitamente atualizar coisa alguma.

Funções personalizadas cancelable permitem com que você cancele a execução de uma função personalizada de streaming para reduzir o consumo de banda larga, memória de trabalho e carregamento de CPU.

Para declarar uma função como streaming ou cancelable, use as marcas de comentário JSDOC `@stream` ou `@cancelable`. 

### <a name="using-an-invocation-parameter"></a>Usando um parâmetro de invocação

O parâmetro `invocation` é o último parâmetro de qualquer função personalizada por padrão. O parâmetro `invocation` fornece um contexto sobre a célula (como o seu endereço) e também permite com que você use os métodos `setResult` e `onCanceled`. Esses métodos definem o que uma função faz quando a função transmite (`setResult`) ou é cancelada (`onCanceled`).

Se você estiver usando o TypeScript, o manipulador de invocações deve ser do tipo `CustomFunctions.StreamingInvocation` ou `CustomFunctions.CancelableInvocation`.

### <a name="streaming-and-cancelable-function-example"></a>Exemplo das funções streaming e cancelable
O exemplo a seguir é uma função personalizada que adiciona um número ao resultado a cada segundo. Observe o seguinte sobre este código:

- O Excel exibe cada valor novo automaticamente usando o método `setResult`.
- O segundo parâmetro de entrada, invocação, não é exibido para os usuários finais no Excel quando eles selecionam a função no menu de preenchimento automático.
- O retorno de chamada `onCanceled` define a função que é executada quando a função é cancelada.

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
> O Excel cancela a execução de uma função nas seguintes situações:
>
> - Quando o usuário edita ou exclui uma célula que faz referência à função.
> - Quando é alterado um dos argumentos (entradas) para a função. Nesse caso, uma nova chamada de função é disparada, seguindo o cancelamento.
> - Quando o usuário aciona manualmente um recálculo. Nesse caso, uma nova chamada de função é disparada, seguindo o cancelamento.

## <a name="next-steps"></a>Próximas etapas

* Saiba mais sobre [diferentes tipos de parâmetros que as suas funções podem usar](custom-functions-parameter-options.md).
* Descubra como [agrupar várias chamadas de API](custom-functions-batching.md).

## <a name="see-also"></a>Confira também

* [Valores voláteis nas funções](custom-functions-volatile.md)
* [Criar metadados JSON para funções personalizadas](custom-functions-json-autogeneration.md)
* [Metadados de funções personalizadas](custom-functions-json.md)
* [Tempo de execução de funções personalizadas do Excel](custom-functions-runtime.md)
* [Práticas recomendadas de funções personalizadas](custom-functions-best-practices.md).
* [Criar funções personalizadas no Excel](custom-functions-overview.md)
* [Tutorial de funções personalizadas do Excel](../tutorials/excel-tutorial-create-custom-functions.md)
