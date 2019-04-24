---
ms.date: 03/21/2019
description: Solicite, transmita e cancele o fluxo de dados externos para sua pasta de trabalho com funções personalizadas no Excel
title: Solicitações da Web e outros dados de tratamento com funções personalizadas (prévia)
localization_priority: Priority
ms.openlocfilehash: 9256e2aa87ec6d7b314314a1e4bc2b3793f1df5c
ms.sourcegitcommit: 9e7b4daa8d76c710b9d9dd4ae2e3c45e8fe07127
ms.translationtype: HT
ms.contentlocale: pt-BR
ms.lasthandoff: 04/24/2019
ms.locfileid: "32449705"
---
# <a name="receiving-and-handling-data-with-custom-functions"></a>Recebimento e gerenciamento de dados com funções personalizadas

Uma das maneiras pelas quais as funções personalizadas aprimoram o poder do Excel é receber dados de locais diferentes na pasta de trabalho, como a web ou um servidor (por meio de WebSockets). As funções personalizadas podem solicitar dados por meio de XHR e buscar solicitações, bem como transmitir esses dados em tempo real.

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

Para outro exemplo de solicitação XHR com mais contexto, confira a função`getFile` dentro [deste arquivo](https://github.com/OfficeDev/Office-Add-in-JavaScript-FileDownload/blob/master/FileDownloadSampleWeb/Home.js) no repositório Github [Office-Add-in-JavaScript-FileDownload](https://github.com/OfficeDev/Office-Add-in-JavaScript-FileDownload).

### <a name="fetch-example"></a>Exemplo de busca

No seguinte exemplo de código, a função stockPriceStream usa um símbolo de cotação da bolsa para acessar o preço de uma ação a cada 1000 milissegundos. Para saber mais sobre este exemplo e obter as JSON acompanhante, confira a [tutorial de funções personalizados](https://docs.microsoft.com/office/dev/add-ins/tutorials/excel-tutorial-create-custom-functions?tabs=excel-windows#create-a-streaming-asynchronous-custom-function). 

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

## <a name="receiving-data-via-websockets"></a>Como receber dados por meio de WebSockets

Em uma função personalizada, é possível usar WebSockets para trocar dados por meio de uma conexão persistente com um servidor. Usando WebSockets, a função personalizada pode abrir uma conexão com um servidor e, em seguida, receber mensagens do servidor automaticamente, quando determinados eventos ocorrerem, sem precisar consultar explicitamente os dados do servidor.

### <a name="websockets-example"></a>Exemplo de WebSockets

O código de exemplo a seguir estabelece uma conexão WebSocket e registra cada mensagem de entrada do servidor.

```JavaScript
var ws = new WebSocket('wss://bundles.office.com');

ws.onmessage(message) {
    console.log(`Recieved: ${message}`);
}

ws.onerror(error){
    console.err(`Failed: ${error}`);
}
```

## <a name="streaming-functions"></a>Funções Streaming

Funções personalizadas de streaming permitem a saída de dados das células repetidamente ao longo do tempo, sem a necessidade de um usuário explicitamente solicitar a atualização de dados. O exemplo a seguir é uma função personalizada que adiciona um número ao resultado a cada segundo. Observe o seguinte sobre este código:

- Cada novo valor usando o Excel automaticamente exibirá o retorno de chamada setResult.
- O segundo parâmetro de entrada, identificador, não é exibido para os usuários finais no Excel quando eles selecionam a função no menu de preenchimento automático.
- O retorno de chamada onCanceled define a função que é executada quando a função é cancelada. Implemente um identificador de cancelamento assim para qualquer função de streaming. Para saber mais, confira [Cancelar uma função](#canceling-a-function).

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

Quando você especifica os metadados para uma função streaming no arquivo de metadados JSON, você deve definir as propriedades “cancelable”. verdadeiro e “stream”. verdadeiro dentro do objeto opções, conforme mostrado no exemplo a seguir. 

```JSON
{
  "id": "INCREMENT",
  "name": "INCREMENT",
  "description": "Periodically increment a value",
  "helpUrl": "http://www.contoso.com",
  "result": {
    "type": "number",
    "dimensionality": "scalar"
  },
  "parameters": [
    {
      "name": "increment",
      "description": "Amount to increment",
      "type": "number",
      "dimensionality": "scalar"
    }
  ],
  "options": {
    "cancelable": true,
    "stream": true
  }
}
```

## <a name="canceling-a-function"></a>Cancelar uma função

Em algumas situações, talvez seja necessário cancelar a execução de uma função personalizada de streaming para reduzir o consumo de banda larga, memória de trabalho e carregamento de CPU. O Excel cancela a execução de uma função nas seguintes situações:

- Quando o usuário edita ou exclui uma célula que faz referência à função.
- Quando é alterado um dos argumentos (entradas) para a função. Nesse caso, uma nova chamada de função é disparada, seguindo o cancelamento.
- Quando o usuário aciona manualmente um recálculo. Nesse caso, uma nova chamada de função é disparada, seguindo o cancelamento.

Para tornar uma função possível de ser cancelada, implemente um identificador de código de função para informar o que fazer quando ela for cancelada. Além disso, especifique a propriedade `"cancelable": true` contida no objeto opções nos metadados JSON que descreve a função. Amostras de código na seção anterior neste artigo fornecem um exemplo dessas técnicas.

## <a name="see-also"></a>Confira também

* [Tutorial de funções personalizadas do Excel](../tutorials/excel-tutorial-create-custom-functions.md)
* [Metadados de funções personalizadas](custom-functions-json.md)
* [Tempo de execução de funções personalizadas do Excel](custom-functions-runtime.md)
* [Práticas recomendadas de funções personalizadas](custom-functions-best-practices.md).
* [Log de alteração de funções personalizadas](custom-functions-changelog.md)
