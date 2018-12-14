---
ms.date: 12/5/2018
description: Entenda os principais cenários de desenvolvimento de funções personalizadas do Excel que usam o novo tempo de execução do JavaScript.
title: Tempo de execução de funções personalizadas do Excel
ms.openlocfilehash: 715690c5cba2466e4a50ba2a33d2324a1abe02f5
ms.sourcegitcommit: 3d8454055ba4d7aae12f335def97357dea5beb30
ms.translationtype: HT
ms.contentlocale: pt-BR
ms.lasthandoff: 12/14/2018
ms.locfileid: "27270828"
---
# <a name="runtime-for-excel-custom-functions-preview"></a><span data-ttu-id="c3964-103">Tempo de execução de funções personalizadas do Excel (versão prévia)</span><span class="sxs-lookup"><span data-stu-id="c3964-103">Runtime for Excel custom functions (preview)</span></span>

<span data-ttu-id="c3964-104">Funções personalizadas usam um novo tempo de execução do JavaScript, diferente do tempo de execução usado por outras partes de um suplemento, como o painel de tarefas ou outros elementos da interface do usuário.</span><span class="sxs-lookup"><span data-stu-id="c3964-104">Custom functions use a new JavaScript runtime that differs from the runtime used by other parts of an add-in, such as the task pane or other UI elements.</span></span> <span data-ttu-id="c3964-105">Esse tempo de execução do JavaScript foi projetado para otimizar o desempenho de cálculos em funções personalizadas, e expõe as novas APIs disponíveis para executar ações comuns baseadas na Web, dentro de funções personalizadas, como solicitação de dados externos ou troca de dados por meio de uma conexão persistente com um servidor.</span><span class="sxs-lookup"><span data-stu-id="c3964-105">This JavaScript runtime is designed to optimize performance of calculations in custom functions and exposes new APIs that you can use to perform common web-based actions within custom functions such as requesting external data or exchanging data over a persistent connection with a server.</span></span> <span data-ttu-id="c3964-106">O tempo de execução do JavaScript também fornece acesso às novas APIs no namespace `OfficeRuntime` que pode ser usado em funções personalizadas ou por outras partes de um suplemento para armazenar dados ou exibir uma caixa de diálogo.</span><span class="sxs-lookup"><span data-stu-id="c3964-106">The JavaScript runtime also provides access to new APIs in the `OfficeRuntime` namespace that can be used within custom functions or by other parts of an add-in to store data or display a dialog box.</span></span> <span data-ttu-id="c3964-107">Este artigo mostra como usar essas APIs em funções personalizadas e descreve considerações adicionais para o desenvolvimento de funções personalizadas.</span><span class="sxs-lookup"><span data-stu-id="c3964-107">This article describes how to use these APIs within custom functions and also outlines additional considerations to keep in mind as you develop custom functions.</span></span>

[!include[Excel custom functions note](../includes/excel-custom-functions-note.md)]

## <a name="requesting-external-data"></a><span data-ttu-id="c3964-108">Como solicitar dados externos</span><span class="sxs-lookup"><span data-stu-id="c3964-108">Requesting external data</span></span>

<span data-ttu-id="c3964-109">É possível solicitar dados externos em uma função personalizada por meio de uma API, como a API [Fetch](https://developer.mozilla.org/en-US/docs/Web/API/Fetch_API), ou por meio de um objeto [XmlHttpRequest (XHR)](https://developer.mozilla.org/en-US/docs/Web/API/XMLHttpRequest), uma API Web padrão que envia solicitações HTTP para interagir com os servidores.</span><span class="sxs-lookup"><span data-stu-id="c3964-109">Within a custom function, you can request external data by using an API like [Fetch](https://developer.mozilla.org/en-US/docs/Web/API/Fetch_API) or by using [XmlHttpRequest (XHR)](https://developer.mozilla.org/en-US/docs/Web/API/XMLHttpRequest), a standard web API that issues HTTP requests to interact with servers.</span></span> <span data-ttu-id="c3964-110">No tempo de execução do JavaScript, o XHR implementa medidas de segurança adicionais solicitando uma [Política de mesma origem](https://developer.mozilla.org/en-US/docs/Web/Security/Same-origin_policy) ou um simples [CORS](https://www.w3.org/TR/cors/).</span><span class="sxs-lookup"><span data-stu-id="c3964-110">Within the JavaScript runtime, XHR implements additional security measures by requiring [Same Origin Policy](https://developer.mozilla.org/en-US/docs/Web/Security/Same-origin_policy) and simple [CORS](https://www.w3.org/TR/cors/).</span></span>  

### <a name="xhr-example"></a><span data-ttu-id="c3964-111">Exemplo de XHR</span><span class="sxs-lookup"><span data-stu-id="c3964-111">XHR example</span></span>

<span data-ttu-id="c3964-112">No código de exemplo a seguir, a função `getTemperature` chama a função `sendWebRequest` para obter a temperatura de uma área específica, de acordo com a ID do termômetro.</span><span class="sxs-lookup"><span data-stu-id="c3964-112">In the following code sample, the `getTemperature` function calls the `sendWebRequest` function to get the temperature of a particular area based on thermometer ID.</span></span> <span data-ttu-id="c3964-113">A função `sendWebRequest` usa XHR para emitir uma solicitação `GET` para um ponto de extremidade que fornece os dados.</span><span class="sxs-lookup"><span data-stu-id="c3964-113">The `sendWebRequest` function uses XHR to issue a `GET` request to an endpoint that can provide the data.</span></span>

> [!NOTE] 
> <span data-ttu-id="c3964-114">Se usar fetch ou XHR, uma nova `Promise` JavaScript será retornada.</span><span class="sxs-lookup"><span data-stu-id="c3964-114">When using fetch or XHR, a new JavaScript `Promise` is returned.</span></span> <span data-ttu-id="c3964-115">Antes de setembro de 2018, era necessário especificar `OfficeExtension.Promise` para usar promessas na API JavaScript para Office, mas agora, basta usar um `Promise` JavaScript.</span><span class="sxs-lookup"><span data-stu-id="c3964-115">Prior to September 2018, you had to specify `OfficeExtension.Promise` to use promises within the Office JavaScript API, but now you can simply use a JavaScript `Promise`.</span></span>

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

## <a name="receiving-data-via-websockets"></a><span data-ttu-id="c3964-116">Como receber dados por meio de WebSockets</span><span class="sxs-lookup"><span data-stu-id="c3964-116">Receiving data via WebSockets</span></span>

<span data-ttu-id="c3964-117">Em uma função personalizada, é possível usar [WebSockets](https://developer.mozilla.org/en-US/docs/Web/API/WebSockets_API) para trocar dados por meio de uma conexão persistente com um servidor.</span><span class="sxs-lookup"><span data-stu-id="c3964-117">Within a custom function, you can use [WebSockets](https://developer.mozilla.org/en-US/docs/Web/API/WebSockets_API) to exchange data over a persistent connection with a server.</span></span> <span data-ttu-id="c3964-118">Usando WebSockets, a função personalizada pode abrir uma conexão com um servidor e, em seguida, receber mensagens do servidor automaticamente, quando determinados eventos ocorrerem, sem precisar consultar explicitamente os dados do servidor.</span><span class="sxs-lookup"><span data-stu-id="c3964-118">By using WebSockets, your custom function can open a connection with a server and then automatically receive messages from the server when certain events occur, without having to explicitly poll the server for data.</span></span>

### <a name="websockets-example"></a><span data-ttu-id="c3964-119">Exemplo de WebSockets</span><span class="sxs-lookup"><span data-stu-id="c3964-119">WebSockets example</span></span>

<span data-ttu-id="c3964-120">O código de exemplo a seguir estabelece uma conexão `WebSocket` e registra cada mensagem de entrada do servidor.</span><span class="sxs-lookup"><span data-stu-id="c3964-120">The following code sample establishes a `WebSocket` connection and then logs each incoming message from the server.</span></span> 

```typescript
const ws = new WebSocket('wss://bundles.office.com');
ws.onmessage = (message) => {
    console.log(`Received: ${message}`);
};
ws.onerror = (error) => {
    console.err(`Failed: ${error}`);
}
```

## <a name="storing-and-accessing-data"></a><span data-ttu-id="c3964-121">Como armazenar e acessar os dados</span><span class="sxs-lookup"><span data-stu-id="c3964-121">Storing and accessing data</span></span>

<span data-ttu-id="c3964-122">Em uma função personalizada (ou em outras partes de um suplemento), você pode armazenar e acessar dados usando o objeto `OfficeRuntime.AsyncStorage`.</span><span class="sxs-lookup"><span data-stu-id="c3964-122">Within a custom function (or within any other part of an add-in), you can store and access data by using the `OfficeRuntime.AsyncStorage` object.</span></span> <span data-ttu-id="c3964-123">`AsyncStorage` é um sistema de armazenamento de chave-valor persistente e não criptografado, que fornece uma alternativa para [localStorage](https://developer.mozilla.org/en-US/docs/Web/API/Window/localStorage), que não pode ser usado em funções personalizadas.</span><span class="sxs-lookup"><span data-stu-id="c3964-123">`AsyncStorage` is a persistent, unencrypted, key-value storage system that provides an alternative to [localStorage](https://developer.mozilla.org/en-US/docs/Web/API/Window/localStorage), which cannot be used within custom functions.</span></span> <span data-ttu-id="c3964-124">Um suplemento pode armazenar até 10 MB de dados por meio de `AsyncStorage`.</span><span class="sxs-lookup"><span data-stu-id="c3964-124">An add-in can store up to 10 MB of data using `AsyncStorage`.</span></span>

<span data-ttu-id="c3964-125">`AsyncStorage` é uma solução de armazenamento compartilhado, o que significa que várias partes de um suplemento podem acessar os mesmos dados.</span><span class="sxs-lookup"><span data-stu-id="c3964-125">`AsyncStorage` is intended as a shared storage solution, meaning multiple parts of an add-in are able to access the same data.</span></span> <span data-ttu-id="c3964-126">Por exemplo, tokens para autenticação de usuário podem ser armazenados em `AsyncStorage`, já que ele pode ser acessado tanto por uma função personalizada quanto por elementos da interface do usuário de um suplemento, como um painel de tarefas.</span><span class="sxs-lookup"><span data-stu-id="c3964-126">For example, tokens for user authentication may be stored in `AsyncStorage` because it can be accessed by both a custom function and add-in UI elements such as a task pane.</span></span> <span data-ttu-id="c3964-127">Da mesma forma, quando dois suplementos compartilham o mesmo domínio (por exemplo, www.contoso.com/suplemento1, www.contoso.com/suplemento2), eles também podem compartilhar informações por meio de `AsyncStorage`.</span><span class="sxs-lookup"><span data-stu-id="c3964-127">Similarly, if two add-ins share the same domain (e.g. www.contoso.com/addin1, www.contoso.com/addin2), they are also permitted to share information back and forth through `AsyncStorage`.</span></span> <span data-ttu-id="c3964-128">Observe que os suplementos que têm diferentes subdomínios terão diferentes instâncias de `AsyncStorage`; por exemplo, subdominio.contoso.com/suplemento1, diferentesubdominio.contoso.com/suplemento2.</span><span class="sxs-lookup"><span data-stu-id="c3964-128">Note that add-ins which have different subdomains will have different instances of `AsyncStorage` (e.g. subdomain.contoso.com/addin1, differentsubdomain.contoso.com/addin2).</span></span> 

<span data-ttu-id="c3964-129">Como `AsyncStorage` pode ser um local compartilhado, é importante notar que é possível substituir os pares chave-valor.</span><span class="sxs-lookup"><span data-stu-id="c3964-129">Because `AsyncStorage` can be a shared location, it is important to realize that it is possible to override key-value pairs.</span></span>

<span data-ttu-id="c3964-130">Os métodos a seguir estão disponíveis no objeto `AsyncStorage`:</span><span class="sxs-lookup"><span data-stu-id="c3964-130">The following methods are available on the `AsyncStorage` object:</span></span>
 
 - `getItem`
 - `setItem`
 - `removeItem`
 - `getAllKeys`
 - `flushGetRequests`
 - `multiGet`
 - `multiSet`
 - <span data-ttu-id="c3964-131">`multiRemove`: você notará que não há implementação de um método para limpar todas as informações (como `clear`).</span><span class="sxs-lookup"><span data-stu-id="c3964-131">`multiRemove`: You will note that there is no implementation of a method for clearing all information (such as `clear`).</span></span> <span data-ttu-id="c3964-132">Em vez disso, use `multiRemove` para remover várias entradas de uma só vez.</span><span class="sxs-lookup"><span data-stu-id="c3964-132">Instead, you should instead use `multiRemove` to remove multiple entries at a time.</span></span>

### <a name="asyncstorage-example"></a><span data-ttu-id="c3964-133">Exemplo de AsyncStorage</span><span class="sxs-lookup"><span data-stu-id="c3964-133">AsyncStorage example</span></span> 

<span data-ttu-id="c3964-134">O exemplo de código a seguir chama a função `AsyncStorage.getItem` para recuperar um valor de armazenamento.</span><span class="sxs-lookup"><span data-stu-id="c3964-134">The following code sample calls the `AsyncStorage.getItem` function to retrieve a value from storage.</span></span>

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

## <a name="displaying-a-dialog-box"></a><span data-ttu-id="c3964-135">Exibindo uma caixa de diálogo</span><span class="sxs-lookup"><span data-stu-id="c3964-135">Displaying a dialog box</span></span>

<span data-ttu-id="c3964-136">Em uma função personalizada (ou em outras partes de um suplemento), você pode usa a API `OfficeRuntime.displayWebDialogOptions` para exibir uma caixa de diálogo.</span><span class="sxs-lookup"><span data-stu-id="c3964-136">Within a custom function (or within any other part of an add-in), you can use the `OfficeRuntime.displayWebDialogOptions` API to display a dialog box.</span></span> <span data-ttu-id="c3964-137">Esta API da caixa de diálogo fornece uma alternativa a [API da caixa de diálogo](../develop/dialog-api-in-office-add-ins.md) que está disponível para uso em painéis de tarefas e comandos de suplemento, mas não em funções personalizadas.</span><span class="sxs-lookup"><span data-stu-id="c3964-137">This dialog API provides an alternative to the [Dialog API](../develop/dialog-api-in-office-add-ins.md) that can be used within task panes and add-in commands, but not within custom functions.</span></span>

### <a name="dialog-api-example"></a><span data-ttu-id="c3964-138">Exemplo de API da caixa de diálogo</span><span class="sxs-lookup"><span data-stu-id="c3964-138">Dialog API example</span></span>

<span data-ttu-id="c3964-139">No exemplo de código a seguir, a função `getTokenViaDialog` usa a função `displayWebDialogOptions` da API da caixa de diálogo para exibir uma caixa de diálogo.</span><span class="sxs-lookup"><span data-stu-id="c3964-139">In the following code sample, the function `getTokenViaDialog` uses the Dialog API’s `displayWebDialogOptions` function to display a dialog box.</span></span>

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

## <a name="additional-considerations"></a><span data-ttu-id="c3964-140">Considerações adicionais</span><span class="sxs-lookup"><span data-stu-id="c3964-140">Additional considerations</span></span>

<span data-ttu-id="c3964-141">Para criar um suplemento que será executado em várias plataformas (um dos principais locatários de Suplementos do Office), você não deve acessar o DOM (Modelo de Objeto do Documento) em funções personalizadas nem usar bibliotecas, como a jQuery, que dependem do DOM.</span><span class="sxs-lookup"><span data-stu-id="c3964-141">In order to create an add-in that will run on multiple platforms (one of the key tenants of Office Add-ins), you should not access the Document Object Model (DOM) in custom functions or use libraries like jQuery that rely on the DOM.</span></span> <span data-ttu-id="c3964-142">No Excel para Windows, onde as funções personalizadas usam o tempo de execução do JavaScript, as funções personalizadas não podem acessar o DOM.</span><span class="sxs-lookup"><span data-stu-id="c3964-142">On Excel for Windows, where custom functions use the JavaScript runtime, custom functions cannot access the DOM.</span></span>

## <a name="see-also"></a><span data-ttu-id="c3964-143">Confira também</span><span class="sxs-lookup"><span data-stu-id="c3964-143">See also</span></span>

* [<span data-ttu-id="c3964-144">Criar funções personalizadas no Excel</span><span class="sxs-lookup"><span data-stu-id="c3964-144">Create custom functions in Excel</span></span>](custom-functions-overview.md)
* [<span data-ttu-id="c3964-145">Metadados de funções personalizadas</span><span class="sxs-lookup"><span data-stu-id="c3964-145">Custom functions metadata</span></span>](custom-functions-json.md)
* <span data-ttu-id="c3964-146">[Práticas recomendadas de funções personalizadas](custom-functions-best-practices.md).</span><span class="sxs-lookup"><span data-stu-id="c3964-146">[Custom functions best practices](custom-functions-best-practices.md)</span></span>
* [<span data-ttu-id="c3964-147">Tutorial de funções personalizadas do Excel</span><span class="sxs-lookup"><span data-stu-id="c3964-147">Excel custom functions tutorial</span></span>](excel-tutorial-custom-functions.md)
