---
ms.date: 10/17/2018
description: Entenda os principais cenários de desenvolvimento de funções personalizadas do Excel que usam o novo tempo de execução do JavaScript.
title: Tempo de execução de funções personalizadas do Excel
ms.openlocfilehash: 7261eb214e97a2a05e08571a31ac9b449df14083
ms.sourcegitcommit: 161a0625646a8c2ebaf1773c6369ee7cc96aa07b
ms.translationtype: HT
ms.contentlocale: pt-BR
ms.lasthandoff: 11/01/2018
ms.locfileid: "25891708"
---
# <a name="runtime-for-excel-custom-functions-preview"></a>Tempo de execução de funções personalizadas do Excel (versão prévia)

Funções personalizadas usam um novo tempo de execução do JavaScript, diferente do tempo de execução usado por outras partes de um suplemento, como o painel de tarefas ou outros elementos da interface do usuário. Esse tempo de execução do JavaScript foi projetado para otimizar o desempenho de cálculos em funções personalizadas, e expõe as novas APIs disponíveis para executar ações comuns baseadas na Web, dentro de funções personalizadas, como solicitação de dados externos ou troca de dados por meio de uma conexão persistente com um servidor. O tempo de execução do JavaScript também fornece acesso às novas APIs no namespace `OfficeRuntime` que pode ser usado em funções personalizadas ou por outras partes de um suplemento para armazenar dados ou exibir uma caixa de diálogo. Este artigo mostra como usar essas APIs em funções personalizadas e descreve considerações adicionais para o desenvolvimento de funções personalizadas.

[!include[Excel custom functions note](../includes/excel-custom-functions-note.md)]

## <a name="requesting-external-data"></a>Como solicitar dados externos

É possível solicitar dados externos em uma função personalizada por meio de uma API, como a API [Fetch](https://developer.mozilla.org/pt-BR/docs/Web/API/Fetch_API), ou por meio de um objeto [XmlHttpRequest (XHR)](https://developer.mozilla.org/pt-BR/docs/Web/API/XMLHttpRequest), uma API Web padrão que envia solicitações HTTP para interagir com os servidores. No tempo de execução do JavaScript, o XHR implementa medidas de segurança adicionais solicitando uma [Política de mesma origem](https://developer.mozilla.org/pt-BR/docs/Web/Security/Same-origin_policy) ou um simples [CORS](https://www.w3.org/TR/cors/).  

### <a name="xhr-example"></a>Exemplo de XHR

No código de exemplo a seguir, a função `getTemperature` chama a função `sendWebRequest` para obter a temperatura de uma área específica, de acordo com a ID do termômetro. A função `sendWebRequest` usa XHR para emitir uma solicitação `GET` para um ponto de extremidade que fornece os dados. 

> [!NOTE] 
> Se usar fetch ou XHR, uma nova `Promise` JavaScript será retornada. Antes de setembro de 2018, era necessário especificar `OfficeExtension.Promise` para usar promessas na API JavaScript para Office, mas agora, basta usar um `Promise` JavaScript.

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

## <a name="receiving-data-via-websockets"></a>Como receber dados por meio de WebSockets

Em uma função personalizada, é possível usar [WebSockets](https://developer.mozilla.org/pt-BR/docs/Web/API/WebSockets_API) para trocar dados por meio de uma conexão persistente com um servidor. Usando WebSockets, a função personalizada pode abrir uma conexão com um servidor e, em seguida, receber mensagens do servidor automaticamente, quando determinados eventos ocorrerem, sem precisar consultar explicitamente os dados do servidor.

### <a name="websockets-example"></a>Exemplo de WebSockets

O código de exemplo a seguir estabelece uma conexão `WebSocket` e registra cada mensagem de entrada do servidor. 

```typescript
const ws = new WebSocket('wss://bundles.office.com');
ws.onmessage = (message) => {
    console.log(`Received: ${message}`);
};
ws.onerror = (error) => {
    console.err(`Failed: ${error}`);
}
```

## <a name="storing-and-accessing-data"></a>Como armazenar e acessar os dados

Em uma função personalizada (ou em outras partes de um suplemento), você pode armazenar e acessar dados usando o objeto `OfficeRuntime.AsyncStorage`. `AsyncStorage` é um sistema de armazenamento de chave-valor persistente e não criptografado, que fornece uma alternativa para [localStorage](https://developer.mozilla.org/pt-BR/docs/Web/API/Window/localStorage), que não pode ser usado em funções personalizadas. Um suplemento pode armazenar até 10 MB de dados por meio de `AsyncStorage`.

`AsyncStorage` é uma solução de armazenamento compartilhado, o que significa que várias partes de um suplemento podem acessar os mesmos dados. Por exemplo, tokens para autenticação de usuário podem ser armazenados em `AsyncStorage`, já que ele pode ser acessado tanto por uma função personalizada quanto por elementos da interface do usuário de um suplemento, como um painel de tarefas. Da mesma forma, quando dois suplementos compartilham o mesmo domínio (por exemplo, www.contoso.com/suplemento1, www.contoso.com/suplemento2), eles também podem compartilhar informações por meio de `AsyncStorage`. Observe que os suplementos que têm diferentes subdomínios terão diferentes instâncias de `AsyncStorage`; por exemplo, subdominio.contoso.com/suplemento1, diferentesubdominio.contoso.com/suplemento2. 

Como `AsyncStorage` pode ser um local compartilhado, é importante notar que é possível substituir os pares chave-valor.

Os métodos a seguir estão disponíveis no objeto `AsyncStorage`:
 
 - `getItem`
 - `setItem`
 - `removeItem`
 - `getAllKeys`
 - `flushGetRequests`
 - `multiGet`
 - `multiSet`
 - `multiRemove`: você notará que não há implementação de um método para limpar todas as informações (como `clear`). Em vez disso, use `multiRemove` para remover várias entradas de uma só vez.

### <a name="asyncstorage-example"></a>Exemplo de AsyncStorage 

O exemplo de código a seguir chama a função `AsyncStorage.getItem` para recuperar um valor de armazenamento.

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

## <a name="displaying-a-dialog-box"></a>Exibindo uma caixa de diálogo

Em uma função personalizada (ou em outras partes de um suplemento), você pode usa a API `OfficeRuntime.displayWebDialogOptions` para exibir uma caixa de diálogo. Esta API da caixa de diálogo fornece uma alternativa a [API da caixa de diálogo](../develop/dialog-api-in-office-add-ins.md) que está disponível para uso em painéis de tarefas e comandos de suplemento, mas não em funções personalizadas.

### <a name="dialog-api-example"></a>Exemplo de API da caixa de diálogo 

No exemplo de código a seguir, a função `getTokenViaDialog` usa a função `displayWebDialogOptions` da API da caixa de diálogo para exibir uma caixa de diálogo.

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

Para criar um suplemento que será executado em várias plataformas (um dos principais locatários de Suplementos do Office), você não deve acessar o DOM (Modelo de Objeto do Documento) em funções personalizadas nem usar bibliotecas, como a jQuery, que dependem do DOM. No Excel para Windows, onde as funções personalizadas usam o tempo de execução do JavaScript, as funções personalizadas não podem acessar o DOM.

## <a name="see-also"></a>Confira também

* [Criar funções personalizadas no Excel](custom-functions-overview.md)
* [Metadados de funções personalizadas](custom-functions-json.md)
* [Práticas recomendadas de funções personalizadas](custom-functions-best-practices.md).
* [Tutorial de funções personalizadas do Excel](excel-tutorial-custom-functions.md)
