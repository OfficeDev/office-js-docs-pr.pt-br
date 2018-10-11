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
# <a name="runtime-for-excel-custom-functions-preview"></a><span data-ttu-id="56e24-103">Runtime de funções personalizadas do Excel (versão prévia)</span><span class="sxs-lookup"><span data-stu-id="56e24-103">Runtime for Excel custom functions</span></span>

<span data-ttu-id="56e24-104">As funções personalizadas usam um novo runtime JavaScript que difere do runtime usado por outras partes de um suplemento, como o painel de tarefas ou outros elementos de interface do usuário.</span><span class="sxs-lookup"><span data-stu-id="56e24-104">Custom functions use a new JavaScript runtime that differs from the runtime used by other parts of an add-in, such as the task pane or other UI elements.</span></span> <span data-ttu-id="56e24-105">Esse runtime Javascript é projetado para otimizar o desempenho dos cálculos em funções personalizadas e expõe novas APIs que você pode usar para executar ações comuns baseadas na Web dentro de funções personalizadas como solicitar dados externos ou trocar dados sobre uma conexão persistente com um servidor.</span><span class="sxs-lookup"><span data-stu-id="56e24-105">This JavaScript runtime is designed to optimize performance of calculations in custom functions and exposes new APIs that you can use to perform common web-based actions within custom functions such as requesting external data or exchanging data over a persistent connection with a server.</span></span> <span data-ttu-id="56e24-106">Esse runtime JavaScript também dá acesso a novas APIs no namespace `OfficeRuntime` que podem ser usadas em funções personalizadas ou por outras partes de um suplemento como armazenamento de dados ou exibição de uma caixa de diálogo.</span><span class="sxs-lookup"><span data-stu-id="56e24-106">The JavaScript runtime also provides access to new APIs in the `OfficeRuntime` namespace that can be used within custom functions or by other parts of an add-in to store data or display a dialog box.</span></span> <span data-ttu-id="56e24-107">Este artigo descreve como usar essas APIs em funções personalizadas e também lista considerações adicionais que você deve ter em mente ao desenvolver funções personalizadas.</span><span class="sxs-lookup"><span data-stu-id="56e24-107">This article describes how to use these APIs within custom functions and also outlines additional considerations to keep in mind as you develop custom functions.</span></span>

[!include[Excel custom functions note](../includes/excel-custom-functions-note.md)]

## <a name="requesting-external-data"></a><span data-ttu-id="56e24-108">Solicitação de dados externos</span><span class="sxs-lookup"><span data-stu-id="56e24-108">Requesting external data</span></span>

<span data-ttu-id="56e24-109">Em uma função personalizada, você pode solicitar dados externos usando uma API como [Fetch](https://developer.mozilla.org/en-US/docs/Web/API/Fetch_API) ou usando [XmlHttpRequest (XHR)](https://developer.mozilla.org/en-US/docs/Web/API/XMLHttpRequest), uma API da web padrão que envia solicitações HTTP para interagir com os servidores.</span><span class="sxs-lookup"><span data-stu-id="56e24-109">Within a custom function, you can request external data by using an API like [Fetch](https://developer.mozilla.org/en-US/docs/Web/API/Fetch_API) or by using [XmlHttpRequest (XHR)](https://developer.mozilla.org/en-US/docs/Web/API/XMLHttpRequest), a standard web API that issues HTTP requests to interact with servers.</span></span> <span data-ttu-id="56e24-110">No novo runtime do JavaScript, XHR implementa medidas adicionais de segurança, exigindo a [Política de mesma origem](https://developer.mozilla.org/en-US/docs/Web/Security/Same-origin_policy) e [CORS](https://www.w3.org/TR/cors/) simples.</span><span class="sxs-lookup"><span data-stu-id="56e24-110">In the new JavaScript runtime, XHR implements additional security measures by requiring [Same Origin Policy](https://developer.mozilla.org/en-US/docs/Web/Security/Same-origin_policy) and simple [CORS](https://www.w3.org/TR/cors/).</span></span>  

### <a name="xhr-example"></a><span data-ttu-id="56e24-111">Exemplo XHR</span><span class="sxs-lookup"><span data-stu-id="56e24-111">XHR example</span></span>

<span data-ttu-id="56e24-112">No exemplo de código a seguir, a função `getTemperature` chama a função `sendWebRequest` para obter a temperatura de uma determinada área com base no ID de termômetro.</span><span class="sxs-lookup"><span data-stu-id="56e24-112">In the following code sample, the  function sends a web request to get the temperature of a particular area based on thermometer ID.</span></span> <span data-ttu-id="56e24-113">A função `sendWebRequest` usa XHR para fazer uma solicitação `GET` para um ponto de extremidade que pode fornecer os dados.</span><span class="sxs-lookup"><span data-stu-id="56e24-113">The `sendWebRequest` function uses XHR to issue a `GET` request to an endpoint that can provide the data.</span></span> 

> [!NOTE] 
> <span data-ttu-id="56e24-114">Ao  efetuar fetch ou usar XHR, um novo `Promise` JavaScript é retornado.</span><span class="sxs-lookup"><span data-stu-id="56e24-114">When using fetch or XHR, a new JavaScript `Promise` is returned.</span></span> <span data-ttu-id="56e24-115">Até antes de setembro de 2018, você tinha que especificar `OfficeExtension.Promise` para usar promessas na API JavaScript do Office, mas agora, pode simplesmente usar um `Promise` JavaScript.</span><span class="sxs-lookup"><span data-stu-id="56e24-115">Prior to September 2018, you had to specify `OfficeExtension.Promise` to use promises within the Office JavaScript API, but now you can simply use a JavaScript `Promise`.</span></span>

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

## <a name="receiving-data-via-websockets"></a><span data-ttu-id="56e24-116">Receber dados via WebSockets</span><span class="sxs-lookup"><span data-stu-id="56e24-116">Receiving data via WebSockets</span></span>

<span data-ttu-id="56e24-117">Dentro de uma função personalizada, você pode usar [WebSockets](https://developer.mozilla.org/en-US/docs/Web/API/WebSockets_API) para trocar dados através de uma conexão persistente com um servidor.</span><span class="sxs-lookup"><span data-stu-id="56e24-117">Within a custom function, you can use [WebSockets](https://developer.mozilla.org/en-US/docs/Web/API/WebSockets_API) to exchange data over a persistent connection with a server.</span></span> <span data-ttu-id="56e24-118">Usando WebSockets, a sua função personalizada por abrir uma conexão com um servidor e receber mensagens automaticamente quando certos eventos ocorrerem, sem precisar explicitamente buscar dados do servidor .</span><span class="sxs-lookup"><span data-stu-id="56e24-118">By using WebSockets, your custom function can open a connection with a server and then automatically receive messages from the server when certain events occur, without having to explicitly poll the server for data.</span></span>

### <a name="websockets-example"></a><span data-ttu-id="56e24-119">Exemplo de WebSockets</span><span class="sxs-lookup"><span data-stu-id="56e24-119">WebSockets example</span></span>

<span data-ttu-id="56e24-120">O exemplo de código a seguir estabelece uma conexão `WebSocket` e, em seguida, registra cada mensagem de entrada vinda do servidor.</span><span class="sxs-lookup"><span data-stu-id="56e24-120">The following code sample establishes a `WebSocket` connection and then logs each incoming message from the server.</span></span> 

```typescript
const ws = new WebSocket('wss://bundles.office.com');
ws.onmessage = (message) => {
    console.log(`Received: ${message}`);
};
ws.onerror = (error) => {
    console.err(`Failed: ${error}`);
}
```

## <a name="storing-and-accessing-data"></a><span data-ttu-id="56e24-121">Armazenamento e acesso a dados</span><span class="sxs-lookup"><span data-stu-id="56e24-121">Storing and accessing data</span></span>

<span data-ttu-id="56e24-122">Em  uma função personalizada (ou em qualquer parte de um suplemento), você pode armazenar e acessar dados usando o objeto `OfficeRuntime.AsyncStorage` .</span><span class="sxs-lookup"><span data-stu-id="56e24-122">Within a custom function (or within any other part of an add-in), you can store and access data by using the `OfficeRuntime.AsyncStorage` object.</span></span> <span data-ttu-id="56e24-123">`AsyncStorage` é um sistema de armazenamento de chave-valor persistente e descriptografados que oferece uma alternativa a [localStorage](https://developer.mozilla.org/en-US/docs/Web/API/Window/localStorage) que não pode ser usado em funções personalizadas.</span><span class="sxs-lookup"><span data-stu-id="56e24-123">`AsyncStorage` is a persistent, unencrypted, key-value storage system that provides an alternative to [localStorage](https://developer.mozilla.org/en-US/docs/Web/API/Window/localStorage), which cannot be used within custom functions.</span></span> <span data-ttu-id="56e24-124">Um suplemento pode armazenar até 10 MB de dados usando `AsyncStorage`.</span><span class="sxs-lookup"><span data-stu-id="56e24-124">An add-in can store up to 10 MB of data using `AsyncStorage`.</span></span>

<span data-ttu-id="56e24-125">Os métodos a seguir estão disponíveis no objeto `AsyncStorage`:</span><span class="sxs-lookup"><span data-stu-id="56e24-125">The following methods are available on the `AsyncStorage` object:</span></span>
 
 - `getItem`
 - `setItem`
 - `removeItem`
 - `clear`
 - `getAllKeys`
 - `flushGetRequests`
 - `multiGet`
 - `multiSet`
 - `multiRemove`

### <a name="asyncstorage-example"></a><span data-ttu-id="56e24-126">Exemplo de AsyncStorage</span><span class="sxs-lookup"><span data-stu-id="56e24-126">AsyncStorage example</span></span> 

<span data-ttu-id="56e24-127">O exemplo de código a seguir chama a função `AsyncStorage.getItem` para recuperar um valor armazenado.</span><span class="sxs-lookup"><span data-stu-id="56e24-127">The following code sample calls the `AsyncStorage.getItem` function to retrieve a value from storage.</span></span>

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

## <a name="displaying-a-dialog-box"></a><span data-ttu-id="56e24-128">Exibição de uma caixa de diálogo</span><span class="sxs-lookup"><span data-stu-id="56e24-128">Open a dialog box</span></span>

<span data-ttu-id="56e24-129">Em uma função personalizada (ou em qualquer parte de um suplemento), você pode usar a API `OfficeRuntime.displayWebDialogOptions` para exibir uma caixa de diálogo.</span><span class="sxs-lookup"><span data-stu-id="56e24-129">Within a custom function (or within any other part of an add-in), you can use the `OfficeRuntime.displayWebDialogOptions` API to display a dialog box.</span></span> <span data-ttu-id="56e24-130">Essa API de caixa de diálogo oferece uma alternativa para a [API de diálogo](../develop/dialog-api-in-office-add-ins.md) que pode ser usada em painéis de tarefas e comandos do suplemento, mas não em funções personalizadas.</span><span class="sxs-lookup"><span data-stu-id="56e24-130">This dialog API provides an alternative to the [Dialog API](../develop/dialog-api-in-office-add-ins.md) that can be used within task panes and add-in commands, but not within custom functions.</span></span>

### <a name="dialog-api-example"></a><span data-ttu-id="56e24-131">Exemplo da API de diálogo</span><span class="sxs-lookup"><span data-stu-id="56e24-131">Dialog API example</span></span> 

<span data-ttu-id="56e24-132">No exemplo de código a seguir, a função `getTokenViaDialog` usa a API de diálogo `displayWebDialogOptions` função para exibir uma caixa de diálogo.</span><span class="sxs-lookup"><span data-stu-id="56e24-132">In the following code sample, the `getTokenViaDialog` method uses the Dialog API’s `displayWebDialogOptions` method to open a dialog box.</span></span>

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

## <a name="additional-considerations"></a><span data-ttu-id="56e24-133">Considerações adicionais</span><span class="sxs-lookup"><span data-stu-id="56e24-133">Additional considerations</span></span>

<span data-ttu-id="56e24-134">Para criar um suplemento que possa ser executado em múltiplas plataformas (um dos locatários principais de Suplementos do Office), você não deve acessar o Document Object Model (DOM) em funções personalizadas ou usar bibliotecas como a jQuery que dependem do DOM.</span><span class="sxs-lookup"><span data-stu-id="56e24-134">In order to create an add-in that will run on multiple platforms (one of the key tenants of Office Add-ins), you should not access the Document Object Model (DOM) in custom functions or use libraries like jQuery that rely on the DOM.</span></span> <span data-ttu-id="56e24-135">No Excel para Windows, onde as funções personalizadas usam o runtime do JavaScript, elas não podem acessar o DOM.</span><span class="sxs-lookup"><span data-stu-id="56e24-135">On Excel for Windows, where custom functions use the JavaScript runtime, custom functions cannot access the DOM.</span></span>

## <a name="see-also"></a><span data-ttu-id="56e24-136">Confira também</span><span class="sxs-lookup"><span data-stu-id="56e24-136">See also</span></span>

* [<span data-ttu-id="56e24-137">Criar funções personalizadas no Excel</span><span class="sxs-lookup"><span data-stu-id="56e24-137">Create custom functions in Excel (Preview)</span></span>](custom-functions-overview.md)
* [<span data-ttu-id="56e24-138">Metadados de funções personalizadas</span><span class="sxs-lookup"><span data-stu-id="56e24-138">Custom functions metadata</span></span>](custom-functions-json.md)
* [<span data-ttu-id="56e24-139">Melhores práticas de funções personalizadas</span><span class="sxs-lookup"><span data-stu-id="56e24-139">Custom functions best practices</span></span>](custom-functions-best-practices.md)
* [<span data-ttu-id="56e24-140">Tutorial de funções personalizadas do Excel</span><span class="sxs-lookup"><span data-stu-id="56e24-140">Excel custom functions tutorial</span></span>](excel-tutorial-custom-functions.md)
