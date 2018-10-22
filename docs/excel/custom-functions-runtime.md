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
# <a name="runtime-for-excel-custom-functions-preview"></a><span data-ttu-id="7ec61-103">Tempo de execução de funções personalizadas do Excel (versão prévia)</span><span class="sxs-lookup"><span data-stu-id="7ec61-103">Runtime for Excel custom functions</span></span>

<span data-ttu-id="7ec61-p101">Funções personalizadas usam um novo tempo de execução do JavaScript que difere do tempo de execução usado por outras partes de um suplemento, como o painel de tarefas ou outros elementos de interface do usuário. Esse tempo de execução do JavaScript foi projetado para otimizar o desempenho dos cálculos em funções personalizadas e expõe novas APIs que você pode usar para executar ações comuns baseadas na web dentro de funções personalizadas, como solicitar dados externos ou troca de dados em uma conexão persistente com um servidor. O tempo de execução do JavaScript também fornece acesso às novas APIs no namespace `OfficeRuntime` que pode ser usado dentro de funções personalizadas ou por outras partes de um suplemento para armazenar dados ou exibir uma caixa de diálogo. Este artigo descreve como usar essas APIs dentro de funções personalizadas e também descreve considerações adicionais para se ter em mente ao desenvolver funções personalizadas.</span><span class="sxs-lookup"><span data-stu-id="7ec61-p101">Custom functions use a new JavaScript runtime that differs from the runtime used by other parts of an add-in, such as the task pane or other UI elements. This JavaScript runtime is designed to optimize performance of calculations in custom functions and exposes new APIs that you can use to perform common web-based actions within custom functions such as requesting external data or exchanging data over a persistent connection with a server. The JavaScript runtime also provides access to new APIs in the `OfficeRuntime` namespace that can be used within custom functions or by other parts of an add-in to store data or display a dialog box. This article describes how to use these APIs within custom functions and also outlines additional considerations to keep in mind as you develop custom functions.</span></span>

[!include[Excel custom functions note](../includes/excel-custom-functions-note.md)]

## <a name="requesting-external-data"></a><span data-ttu-id="7ec61-108">Solicitação de dados externos</span><span class="sxs-lookup"><span data-stu-id="7ec61-108">Requesting external data</span></span>

<span data-ttu-id="7ec61-p102">Dentro de uma função personalizada, você poderá solicitar dados externos usando uma API como [Buscar](https://developer.mozilla.org/en-US/docs/Web/API/Fetch_API) ou usando [XmlHttpRequest (XHR)](https://developer.mozilla.org/en-US/docs/Web/API/XMLHttpRequest), uma API da web padrão que emite solicitações HTTP para interagir com os servidores. Dentro do tempo de execução do JavaScript, XHR implementa medidas de segurança adicionais, exigindo a [Diretiva de mesma origem](https://developer.mozilla.org/en-US/docs/Web/Security/Same-origin_policy) e [CORS](https://www.w3.org/TR/cors/) simples.</span><span class="sxs-lookup"><span data-stu-id="7ec61-p102">Within a custom function, you can request external data by using an API like [Fetch](https://developer.mozilla.org/en-US/docs/Web/API/Fetch_API) or by using [XmlHttpRequest (XHR)](https://developer.mozilla.org/en-US/docs/Web/API/XMLHttpRequest), a standard web API that issues HTTP requests to interact with servers. Within the JavaScript runtime, XHR implements additional security measures by requiring [Same Origin Policy](https://developer.mozilla.org/en-US/docs/Web/Security/Same-origin_policy) and simple [CORS](https://www.w3.org/TR/cors/).</span></span>  

### <a name="xhr-example"></a><span data-ttu-id="7ec61-111">Exemplo XHR</span><span class="sxs-lookup"><span data-stu-id="7ec61-111">XHR example</span></span>

<span data-ttu-id="7ec61-p103">No exemplo de código a seguir, a função `getTemperature` chama a função `sendWebRequest` para obter a temperatura de uma determinada área com base na ID de termômetro. A função `sendWebRequest` usa XHR para emitir uma solicitação `GET` para um ponto de extremidade que pode fornecer os dados.</span><span class="sxs-lookup"><span data-stu-id="7ec61-p103">In the following code sample, the `getTemperature` function calls the `sendWebRequest` function to get the temperature of a particular area based on thermometer ID. The `sendWebRequest` function uses XHR to issue a `GET` request to an endpoint that can provide the data.</span></span> 

> [!NOTE] 
> <span data-ttu-id="7ec61-p104">Ao usar a busca ou XHR, um novo `Promise` JavaScript é retornado. Antes de setembro de 2018, era necessário especificar `OfficeExtension.Promise` para usar promessas dentro da API JavaScript do Office, mas agora você pode simplesmente usar um JavaScript `Promise`.</span><span class="sxs-lookup"><span data-stu-id="7ec61-p104">When using fetch or XHR, a new JavaScript `Promise` is returned. Prior to September 2018, you had to specify `OfficeExtension.Promise` to use promises within the Office JavaScript API, but now you can simply use a JavaScript `Promise`.</span></span>

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

## <a name="receiving-data-via-websockets"></a><span data-ttu-id="7ec61-116">Receber dados via WebSockets</span><span class="sxs-lookup"><span data-stu-id="7ec61-116">Receiving data via WebSockets</span></span>

<span data-ttu-id="7ec61-p105">Dentro de uma função personalizada, você pode usar [WebSockets](https://developer.mozilla.org/en-US/docs/Web/API/WebSockets_API) para trocar dados em uma conexão persistente com um servidor. Usando  WebSockets, sua função personalizada poderá abrir uma conexão com um servidor e, em seguida, automaticamente receber mensagens do servidor quando determinados eventos ocorrerem, sem precisar explicitamente sondar o servidor de dados.</span><span class="sxs-lookup"><span data-stu-id="7ec61-p105">Within a custom function, you can use [WebSockets](https://developer.mozilla.org/en-US/docs/Web/API/WebSockets_API) to exchange data over a persistent connection with a server. By using WebSockets, your custom function can open a connection with a server and then automatically receive messages from the server when certain events occur, without having to explicitly poll the server for data.</span></span>

### <a name="websockets-example"></a><span data-ttu-id="7ec61-119">Exemplo de WebSockets</span><span class="sxs-lookup"><span data-stu-id="7ec61-119">WebSockets example</span></span>

<span data-ttu-id="7ec61-120">O exemplo de código a seguir estabelece uma conexão `WebSocket` e, em seguida, registra cada mensagem de entrada vinda do servidor.</span><span class="sxs-lookup"><span data-stu-id="7ec61-120">The following code sample establishes a `WebSocket` connection and then logs each incoming message from the server.</span></span> 

```typescript
const ws = new WebSocket('wss://bundles.office.com');
ws.onmessage = (message) => {
    console.log(`Received: ${message}`);
};
ws.onerror = (error) => {
    console.err(`Failed: ${error}`);
}
```

## <a name="storing-and-accessing-data"></a><span data-ttu-id="7ec61-121">Armazenamento e acesso a dados</span><span class="sxs-lookup"><span data-stu-id="7ec61-121">Storing and accessing data</span></span>

<span data-ttu-id="7ec61-p106">Dentro de uma função personalizada (ou em qualquer parte de um suplemento), você pode armazenar e acessar dados usando o objeto `OfficeRuntime.AsyncStorage`. `AsyncStorage` é um sistema de armazenamento persistente, não criptografado e de chave-valor que fornece uma alternativa ao [localStorage](https://developer.mozilla.org/en-US/docs/Web/API/Window/localStorage), que não pode ser usado dentro de funções personalizadas. Um suplemento pode armazenar até 10 MB de dados usando `AsyncStorage`.</span><span class="sxs-lookup"><span data-stu-id="7ec61-p106">Within a custom function (or within any other part of an add-in), you can store and access data by using the `OfficeRuntime.AsyncStorage` object. `AsyncStorage` is a persistent, unencrypted, key-value storage system that provides an alternative to [localStorage](https://developer.mozilla.org/en-US/docs/Web/API/Window/localStorage), which cannot be used within custom functions. An add-in can store up to 10 MB of data using `AsyncStorage`.</span></span>

<span data-ttu-id="7ec61-125">Os métodos a seguir estão disponíveis no objeto `AsyncStorage` :</span><span class="sxs-lookup"><span data-stu-id="7ec61-125">The following methods are available on the `AsyncStorage` object:</span></span>
 
 - `getItem`
 - `setItem`
 - `removeItem`
 - `clear`
 - `getAllKeys`
 - `flushGetRequests`
 - `multiGet`
 - `multiSet`
 - `multiRemove`

### <a name="asyncstorage-example"></a><span data-ttu-id="7ec61-126">Exemplo de AsyncStorage</span><span class="sxs-lookup"><span data-stu-id="7ec61-126">AsyncStorage example</span></span> 

<span data-ttu-id="7ec61-127">O exemplo de código a seguir chama a função `AsyncStorage.getItem` para recuperar um valor armazenado.</span><span class="sxs-lookup"><span data-stu-id="7ec61-127">The following code sample calls the `AsyncStorage.getItem` function to retrieve a value from storage.</span></span>

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

## <a name="displaying-a-dialog-box"></a><span data-ttu-id="7ec61-128">Exibição de uma caixa de diálogo</span><span class="sxs-lookup"><span data-stu-id="7ec61-128">Open a dialog box.</span></span>

<span data-ttu-id="7ec61-p107">Dentro de uma função personalizada (ou em qualquer parte de um suplemento), você pode usar a API `OfficeRuntime.displayWebDialogOptions` para exibir uma caixa de diálogo. Essa API de diálogo oferece uma alternativa para a [API de diálogo](../develop/dialog-api-in-office-add-ins.md) que pode ser usada dentro painéis de tarefas e comandos de suplemento, mas não dentro de funções personalizadas.</span><span class="sxs-lookup"><span data-stu-id="7ec61-p107">Within a custom function (or within any other part of an add-in), you can use the `OfficeRuntime.displayWebDialogOptions` API to display a dialog box. This dialog API provides an alternative to the [Dialog API](../develop/dialog-api-in-office-add-ins.md) that can be used within task panes and add-in commands, but not within custom functions.</span></span>

### <a name="dialog-api-example"></a><span data-ttu-id="7ec61-131">Exemplo da API de diálogo</span><span class="sxs-lookup"><span data-stu-id="7ec61-131">Dialog API example</span></span> 

<span data-ttu-id="7ec61-132">No exemplo de código a seguir, a função `getTokenViaDialog` usa a API de diálogo `displayWebDialogOptions` função para exibir uma caixa de diálogo.</span><span class="sxs-lookup"><span data-stu-id="7ec61-132">In the following code sample, the `getTokenViaDialog` method uses the Dialog API’s `displayWebDialogOptions` method to open a dialog box.</span></span>

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

## <a name="additional-considerations"></a><span data-ttu-id="7ec61-133">Considerações adicionais</span><span class="sxs-lookup"><span data-stu-id="7ec61-133">Additional considerations</span></span>

<span data-ttu-id="7ec61-p108">Para criar um suplemento que será executado em várias plataformas (um dos locatários principais de suplementos do Office), você não deve acessar o modelo de objeto de documento (DOM) em funções personalizadas ou usar bibliotecas como jQuery que dependem de DOM. No Excel para Windows, onde as funções personalizadas usam o tempo de execução do JavaScript, funções personalizadas não podem acessar o DOM.</span><span class="sxs-lookup"><span data-stu-id="7ec61-p108">In order to create an add-in that will run on multiple platforms (one of the key tenants of Office Add-ins), you should not access the Document Object Model (DOM) in custom functions or use libraries like jQuery that rely on the DOM. On Excel for Windows, where custom functions use the JavaScript runtime, custom functions cannot access the DOM.</span></span>

## <a name="see-also"></a><span data-ttu-id="7ec61-136">Confira também</span><span class="sxs-lookup"><span data-stu-id="7ec61-136">See also</span></span>

* [<span data-ttu-id="7ec61-137">Criar funções personalizadas no Excel</span><span class="sxs-lookup"><span data-stu-id="7ec61-137">Create custom functions in Excel (Preview)</span></span>](custom-functions-overview.md)
* [<span data-ttu-id="7ec61-138">Metadados de funções personalizadas</span><span class="sxs-lookup"><span data-stu-id="7ec61-138">Custom functions metadata</span></span>](custom-functions-json.md)
* [<span data-ttu-id="7ec61-139">Melhores práticas de funções personalizadas</span><span class="sxs-lookup"><span data-stu-id="7ec61-139">Custom functions best practices</span></span>](custom-functions-best-practices.md)
* [<span data-ttu-id="7ec61-140">Tutorial de funções personalizadas do Excel</span><span class="sxs-lookup"><span data-stu-id="7ec61-140">Excel custom functions tutorial</span></span>](excel-tutorial-custom-functions.md)
