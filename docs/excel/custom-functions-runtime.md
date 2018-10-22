---
ms.date: 10/17/2018
description: Compreenda os principais cenários no desenvolvimento de funções personalizadas do Excel que usam o novo runtime do JavaScript.
title: Runtime de funções personalizadas do Excel
ms.openlocfilehash: 333816c3916af1490d14b8344c4bb49094f9a7f9
ms.sourcegitcommit: a6d6348075c1abed76d2146ddfc099b0151fe403
ms.translationtype: HT
ms.contentlocale: pt-BR
ms.lasthandoff: 10/19/2018
ms.locfileid: "25640012"
---
# <a name="runtime-for-excel-custom-functions-preview"></a>Tempo de execução de funções personalizadas do Excel (versão prévia)

Funções personalizadas usam um novo tempo de execução do JavaScript que difere do tempo de execução usado por outras partes de um suplemento, como o painel de tarefas ou outros elementos de interface do usuário. Esse tempo de execução do JavaScript foi projetado para otimizar o desempenho dos cálculos em funções personalizadas e expõe novas APIs que você pode usar para executar ações comuns baseadas na web dentro de funções personalizadas, como solicitar dados externos ou troca de dados em uma conexão persistente com um servidor. O tempo de execução do JavaScript também fornece acesso às novas APIs no namespace `OfficeRuntime` que pode ser usado dentro de funções personalizadas ou por outras partes de um suplemento para armazenar dados ou exibir uma caixa de diálogo. Este artigo descreve como usar essas APIs dentro de funções personalizadas e também descreve considerações adicionais para se ter em mente ao desenvolver funções personalizadas.

[!include[Excel custom functions note](../includes/excel-custom-functions-note.md)]

## <a name="requesting-external-data"></a>Solicitação de dados externos

Dentro de uma função personalizada, você poderá solicitar dados externos usando uma API como [Buscar](https://developer.mozilla.org/en-US/docs/Web/API/Fetch_API) ou usando [XmlHttpRequest (XHR)](https://developer.mozilla.org/en-US/docs/Web/API/XMLHttpRequest), uma API da web padrão que emite solicitações HTTP para interagir com os servidores. Dentro do tempo de execução do JavaScript, XHR implementa medidas de segurança adicionais, exigindo a [Diretiva de mesma origem](https://developer.mozilla.org/en-US/docs/Web/Security/Same-origin_policy) e [CORS](https://www.w3.org/TR/cors/) simples.  

### <a name="xhr-example"></a>Exemplo XHR

No exemplo de código a seguir, a função `getTemperature` chama a função `sendWebRequest` para obter a temperatura de uma determinada área com base na ID de termômetro. A função `sendWebRequest` usa XHR para emitir uma solicitação `GET` para um ponto de extremidade que pode fornecer os dados. 

> [!NOTE] 
> Ao usar a busca ou XHR, um novo `Promise` JavaScript é retornado. Antes de setembro de 2018, era necessário especificar `OfficeExtension.Promise` para usar promessas dentro da API JavaScript do Office, mas agora você pode simplesmente usar um JavaScript `Promise`.

```js
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
        xhttp.open("GET", "https://contoso.com/temperature/" + thermometerID), true)
        xhttp.send();  
    }
}
```

## <a name="receiving-data-via-websockets"></a>Receber dados via WebSockets

Dentro de uma função personalizada, você pode usar [WebSockets](https://developer.mozilla.org/en-US/docs/Web/API/WebSockets_API) para trocar dados em uma conexão persistente com um servidor. Usando  WebSockets, sua função personalizada poderá abrir uma conexão com um servidor e, em seguida, automaticamente receber mensagens do servidor quando determinados eventos ocorrerem, sem precisar explicitamente sondar o servidor de dados.

### <a name="websockets-example"></a>Exemplo de WebSockets

O exemplo de código a seguir estabelece uma conexão `WebSocket` e, em seguida, registra cada mensagem de entrada vinda do servidor. 

```typescript
const ws = new WebSocket('wss://bundles.office.com');
ws.onmessage = (message) => {
    console.log(`Received: ${message}`);
};
ws.onerror = (error) => {
    console.err(`Failed: ${error}`);
}
```

## <a name="storing-and-accessing-data"></a>Armazenamento e acesso a dados

Dentro de uma função personalizada (ou em qualquer parte de um suplemento), você pode armazenar e acessar dados usando o objeto `OfficeRuntime.AsyncStorage`. `AsyncStorage` é um sistema de armazenamento persistente, não criptografado e de chave-valor que fornece uma alternativa ao [localStorage](https://developer.mozilla.org/en-US/docs/Web/API/Window/localStorage), que não pode ser usado dentro de funções personalizadas. Um suplemento pode armazenar até 10 MB de dados usando `AsyncStorage`.

Os métodos a seguir estão disponíveis no objeto `AsyncStorage` :
 
 - `getItem`
 - `setItem`
 - `removeItem`
 - `clear`
 - `getAllKeys`
 - `flushGetRequests`
 - `multiGet`
 - `multiSet`
 - `multiRemove`

### <a name="asyncstorage-example"></a>Exemplo de AsyncStorage 

O exemplo de código a seguir chama a função `AsyncStorage.getItem` para recuperar um valor armazenado.

```typescript
_goGetData = async () => {
    try {
        const value = await AsyncStorage.getItem('toDoItem');
        if (value !== null) {
            //data exists and you can do something with it here
        }
    } catch (error) {
        //handle errors here
    }
}
```

## <a name="displaying-a-dialog-box"></a>Exibição de uma caixa de diálogo

Dentro de uma função personalizada (ou em qualquer parte de um suplemento), você pode usar a API `OfficeRuntime.displayWebDialogOptions` para exibir uma caixa de diálogo. Essa API de diálogo oferece uma alternativa para a [API de diálogo](../develop/dialog-api-in-office-add-ins.md) que pode ser usada dentro painéis de tarefas e comandos de suplemento, mas não dentro de funções personalizadas.

### <a name="dialog-api-example"></a>Exemplo da API de diálogo 

No exemplo de código a seguir, a função `getTokenViaDialog` usa a API de diálogo `displayWebDialogOptions` função para exibir uma caixa de diálogo.

```js
// Get auth token before calling my service, a hypothetical API that will deliver a stock price based on stock ticker string, such as "MSFT"

function getStock (ticker) {
  return new Promise(function (resolve, reject) {
    // Get a token
    getToken("https://www.contoso.com/auth")
    .then(function (token) {

      // Use token to get stock price
      fetch("https://www.contoso.com/?token=token&ticker= + ticker")
      .then(function (result) {

        // Return stock price to cell
        resolve(result);
      });
    })
    .catch(function (error) {
      reject(error);
    });
  });
  
  //Helper
  function getToken(url) {
    return new Promise(function (resolve,reject) {
      if(_cachedToken) {
        resolve(_cachedToken);
      } else {
        getTokenViaDialog(url)
        .then(function (result) {
          resolve(result);
        })
        .catch(function (result) {
          reject(result);
        });
      }
    });
  }

  function getTokenViaDialog(url) {
    return new Promise (function (resolve, reject) {
      if (_dialogOpen) {
        // Can only have one dialog open at once, wait for previous dialog's token
        let timeout = 5;
        let count = 0;
        var intervalId = setInterval(function () {
          count++;
          if(_cachedToken) {
            resolve(_cachedToken);
            clearInterval(intervalId);
          }
          if(count >= timeout) {
            reject("Timeout while waiting for token");
            clearInterval(intervalId);
          }
        }, 1000);
      } else {
        _dialogOpen = true;
        OfficeRuntime.displayWebDialogOptions(url, {
          height: '50%',
          width: '50%',
          onMessage: function (message, dialog) {
            _cachedToken = message;
            resolve(message);
            dialog.closeDialog();
            return;
          },
          onRuntimeError: function(error, dialog) {
            reject(error);
          },
        }).catch(function (e) {
          reject(e);
        });
      }
    });
  }
}
```

## <a name="additional-considerations"></a>Considerações adicionais

Para criar um suplemento que será executado em várias plataformas (um dos locatários principais de suplementos do Office), você não deve acessar o modelo de objeto de documento (DOM) em funções personalizadas ou usar bibliotecas como jQuery que dependem de DOM. No Excel para Windows, onde as funções personalizadas usam o tempo de execução do JavaScript, funções personalizadas não podem acessar o DOM.

## <a name="see-also"></a>Confira também

* [Criar funções personalizadas no Excel](custom-functions-overview.md)
* [Metadados de funções personalizadas](custom-functions-json.md)
* [Melhores práticas de funções personalizadas](custom-functions-best-practices.md)
* [Tutorial de funções personalizadas do Excel](excel-tutorial-custom-functions.md)
