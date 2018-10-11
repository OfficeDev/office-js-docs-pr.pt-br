---
ms.date: 10/03/2018
description: Compreenda os principais cenários no desenvolvimento de funções personalizadas do Excel que usam o novo runtime do JavaScript.
title: Runtime de funções personalizadas do Excel
ms.openlocfilehash: a48b02a8ca404b51740d9052d199da934eb9312e
ms.sourcegitcommit: 563c53bac52b31277ab935f30af648f17c5ed1e2
ms.translationtype: HT
ms.contentlocale: pt-BR
ms.lasthandoff: 10/10/2018
ms.locfileid: "25459102"
---
# <a name="runtime-for-excel-custom-functions-preview"></a>Runtime de funções personalizadas do Excel (versão prévia)

As funções personalizadas usam um novo runtime JavaScript que difere do runtime usado por outras partes de um suplemento, como o painel de tarefas ou outros elementos de interface do usuário. Esse runtime Javascript é projetado para otimizar o desempenho dos cálculos em funções personalizadas e expõe novas APIs que você pode usar para executar ações comuns baseadas na Web dentro de funções personalizadas como solicitar dados externos ou trocar dados sobre uma conexão persistente com um servidor. Esse runtime JavaScript também dá acesso a novas APIs no namespace `OfficeRuntime` que podem ser usadas em funções personalizadas ou por outras partes de um suplemento como armazenamento de dados ou exibição de uma caixa de diálogo. Este artigo descreve como usar essas APIs em funções personalizadas e também lista considerações adicionais que você deve ter em mente ao desenvolver funções personalizadas.

[!include[Excel custom functions note](../includes/excel-custom-functions-note.md)]

## <a name="requesting-external-data"></a>Solicitação de dados externos

Em uma função personalizada, você pode solicitar dados externos usando uma API como [Fetch](https://developer.mozilla.org/en-US/docs/Web/API/Fetch_API) ou usando [XmlHttpRequest (XHR)](https://developer.mozilla.org/en-US/docs/Web/API/XMLHttpRequest), uma API da web padrão que envia solicitações HTTP para interagir com os servidores. No novo runtime do JavaScript, XHR implementa medidas adicionais de segurança, exigindo a [Política de mesma origem](https://developer.mozilla.org/en-US/docs/Web/Security/Same-origin_policy) e [CORS](https://www.w3.org/TR/cors/) simples.  

### <a name="xhr-example"></a>Exemplo XHR

No exemplo de código a seguir, a função `getTemperature` chama a função `sendWebRequest` para obter a temperatura de uma determinada área com base no ID de termômetro. A função `sendWebRequest` usa XHR para fazer uma solicitação `GET` para um ponto de extremidade que pode fornecer os dados. 

> [!NOTE] 
> Ao  efetuar fetch ou usar XHR, um novo `Promise` JavaScript é retornado. Até antes de setembro de 2018, você tinha que especificar `OfficeExtension.Promise` para usar promessas na API JavaScript do Office, mas agora, pode simplesmente usar um `Promise` JavaScript.

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

Dentro de uma função personalizada, você pode usar [WebSockets](https://developer.mozilla.org/en-US/docs/Web/API/WebSockets_API) para trocar dados através de uma conexão persistente com um servidor. Usando WebSockets, a sua função personalizada por abrir uma conexão com um servidor e receber mensagens automaticamente quando certos eventos ocorrerem, sem precisar explicitamente buscar dados do servidor .

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

Em  uma função personalizada (ou em qualquer parte de um suplemento), você pode armazenar e acessar dados usando o objeto `OfficeRuntime.AsyncStorage` . `AsyncStorage` é um sistema de armazenamento de chave-valor persistente e descriptografados que oferece uma alternativa a [localStorage](https://developer.mozilla.org/en-US/docs/Web/API/Window/localStorage) que não pode ser usado em funções personalizadas. Um suplemento pode armazenar até 10 MB de dados usando `AsyncStorage`.

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

Em uma função personalizada (ou em qualquer parte de um suplemento), você pode usar a API `OfficeRuntime.displayWebDialogOptions` para exibir uma caixa de diálogo. Essa API de caixa de diálogo oferece uma alternativa para a [API de diálogo](../develop/dialog-api-in-office-add-ins.md) que pode ser usada em painéis de tarefas e comandos do suplemento, mas não em funções personalizadas.

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

Para criar um suplemento que possa ser executado em múltiplas plataformas (um dos locatários principais de Suplementos do Office), você não deve acessar o Document Object Model (DOM) em funções personalizadas ou usar bibliotecas como a jQuery que dependem do DOM. No Excel para Windows, onde as funções personalizadas usam o runtime do JavaScript, elas não podem acessar o DOM.

## <a name="see-also"></a>Confira também

* [Criar funções personalizadas no Excel](custom-functions-overview.md)
* [Metadados de funções personalizadas](custom-functions-json.md)
* [Melhores práticas de funções personalizadas](custom-functions-best-practices.md)
* [Tutorial de funções personalizadas do Excel](excel-tutorial-custom-functions.md)
