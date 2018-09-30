---
ms.date: 09/27/2018
description: Funções personalizadas do Excel usam um novo tempo de execução do JavaScript que difere do tempo de execução de controle do modo de exibição da Web para suplementos padrão.
title: Tempo de execução de funções personalizadas do Excel
ms.openlocfilehash: 7489cd66851d1e0c24ef573ffa920b794cf749c2
ms.sourcegitcommit: 1852ae367de53deb91d03ca55d16eb69709340d3
ms.translationtype: HT
ms.contentlocale: pt-BR
ms.lasthandoff: 09/29/2018
ms.locfileid: "25348756"
---
# <a name="runtime-for-excel-custom-functions-preview"></a>Tempo de execução de funções personalizadas do Excel (versão prévia)

As funções personalizadas estendem as funcionalidades do Excel usando um novo tempo de execução do JavaScript que usa um mecanismo de JavaScript em área restrita em vez de um navegador da web. Como as funções personalizadas não precisam renderizar elementos de interface do usuário, o novo tempo de execução do JavaScript é otimizado para fazer cálculos, permitindo que você execute milhares de funções personalizadas simultaneamente.

[!include[Excel custom functions note](../includes/excel-custom-functions-note.md)]

## <a name="key-facts-about-the-new-javascript-runtime"></a>Fatos importantes sobre o novo tempo de execução do JavaScript 

Somente as funções personalizadas de um suplemento usam o novo tempo de execução do JavaScript descrito neste artigo. Se um suplemento incluir outros componentes, como painéis de tarefas e outros elementos de interface do usuário, além das funções personalizadas, esses outros componentes do suplemento continuarão operando no tempo de execução de exibição da Web com aparência de navegador.  Além disso: 

- O tempo de execução do JavaScript não fornece acesso ao Document Object Model (DOM) ou a bibliotecas de suporte, como jQuery, que dependem do DOM.

- Uma função personalizada que é definida em um arquivo JavaScript de um suplemento pode retornar um `Promise` regular do JavaScript em vez de retornar `OfficeExtension.Promise`.  

- O arquivo JSON que especifica a função personalizada metatdata não precisa especificar **sync** ou **async** nas **opções**.

## <a name="new-apis"></a>Novas APIs 

O tempo de execução do JavaScript que é usado pelas funções personalizadas tem as seguintes APIs:

- [XHR](#xhr)
- [WebSockets](#websockets)
- [AsyncStorage](#asyncstorage)
- [API de diálogo](#dialog-api)

### <a name="xhr"></a>XHR

XHR significa [XmlHttpRequest](https://developer.mozilla.org/en-US/docs/Web/API/XMLHttpRequest), uma API da Web padrão que emite solicitações HTTP para interagir com os servidores. No novo tempo de execução do JavaScript, XHR implementa medidas adicionais de segurança, exigindo a [Mesma diretiva de origem](https://developer.mozilla.org/en-US/docs/Web/Security/Same-origin_policy) e [CORS](https://www.w3.org/TR/cors/) simples.  

No exemplo de código a seguir, a função `getTemperature()` envia uma solicitação da web para obter a temperatura de uma determinada área com base na ID de termômetro. A função `sendWebRequest()` usa XHR para emitir uma solicitação `GET` para um ponto de extremidade que pode fornecer os dados.  

```js
function getTemperature(thermometerID) {
  return new Promise(function(setResult) {
      sendWebRequest(thermometerID, function(data){ //sendWebRequest is defined later in this code sample
          storeLastTemperature(thermometerID, data.temperature);
          setResult(data.temperature);
      });
  });
}

//Helper method that uses Office's implementation of XMLHttpRequest in the new JavaScript runtime for custom functions  
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

### <a name="websockets"></a>WebSockets

[WebSockets](https://developer.mozilla.org/en-US/docs/Web/API/WebSockets_API) é um protocolo de rede que cria a comunicação em tempo real entre um servidor e um ou mais clientes. Ele é frequentemente usado para aplicativos de bate-papo porque permite que você possa ler e gravar texto simultaneamente.  

Como mostra o exemplo de código a seguir, as funções personalizadas podem usar WebSockets. Neste exemplo, o WebSocket registra cada mensagem que recebe.

```typescript
const ws = new WebSocket('wss://bundles.office.com');
ws.onmessage = (message) => {
    console.log(`Received: ${message}`);
};
ws.onerror = (error) => {
    console.err(`Failed: ${error}`);
}
```

### <a name="asyncstorage"></a>AsyncStorage

AsyncStorage é um sistema de armazenamento de chave-valor que pode ser usado para armazenar os tokens de autenticação. Ele é:

- Persistente
- Não encriptado
- Assíncrono

AsyncStorage fica disponível globalmente para todas as partes do seu suplemento. Para funções personalizadas, `AsyncStorage` é exposto como um objeto global. (Para outras partes do seu suplemento, como painéis de tarefas e outros elementos que usam o tempo de execução de exibição da Web, AsyncStorage é exposto por meio do `OfficeRuntime`.) Cada suplemento tem sua própria partição de armazenamento, com um tamanho padrão de 5 MB. 

Os métodos a seguir estão disponíveis no objeto `AsyncStorage`:
 
 - `getItem`
 - `setItem`
 - `removeItem`
 - `clear`
 - `getAllKeys`
 - `flushGetRequests`
 - `multiGet`
 - `multiSet`
 - `multiRemove`
 
Neste momento, os métodos `mergeItem` e `multiMerge` não são suportados.

O seguinte código de amostra chama a função `AsyncStorage.getItem` para recuperar um valor de armazenamento.

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
}
```

### <a name="dialog-api"></a>API de diálogo

A API de diálogo permite que você abra uma caixa de diálogo que solicita a entrada do usuário. Você pode usar a API de diálogo para exigir a autenticação de usuário por meio de um recurso externo, como Google ou Facebook, para que o usuário possa usar sua função.   

No exemplo de código a seguir, o método `getTokenViaDialog()` usa o método `displayWebDialog()` da API de diálogo para abrir uma caixa de diálogo.

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
        OfficeRuntime.displayWebDialog(url, {
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

> [!NOTE]
> A API de diálogo descrita nesta seção faz parte do novo tempo de execução do JavaScript para funções personalizadas e pode ser usada somente nas funções personalizadas. Essa API é diferente da [API de diálogo](../develop/dialog-api-in-office-add-ins.md) que pode ser usada nos painéis de tarefas e comandos do suplemento.

## <a name="see-also"></a>Confira também

* [Criar funções personalizadas no Excel](custom-functions-overview.md)
* [Metadados de funções personalizadas](custom-functions-json.md)
* [Melhores práticas de funções personalizadas](custom-functions-best-practices.md)
* [Tutorial de funções personalizadas do Excel](excel-tutorial-custom-functions.md)
