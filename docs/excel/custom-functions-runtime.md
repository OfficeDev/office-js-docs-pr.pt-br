---
ms.date: 09/20/2018
description: Funções personalizadas do Excel usam um novo tempo de execução do JavaScript que difere do tempo de execução de controle do modo de exibição da Web para suplementos padrão.
title: Tempo de execução de funções personalizados do Excel
ms.openlocfilehash: d31002096fccd682c0f2a23a8b43249af5d4df8f
ms.sourcegitcommit: 470d8212b256275587e651abaa6f28beafebcab4
ms.translationtype: HT
ms.contentlocale: pt-BR
ms.lasthandoff: 09/21/2018
ms.locfileid: "24068808"
---
# <a name="runtime-for-excel-custom-functions"></a><span data-ttu-id="57e9b-103">Tempo de execução de funções personalizados do Excel</span><span class="sxs-lookup"><span data-stu-id="57e9b-103">Runtime for Excel custom functions</span></span>

<span data-ttu-id="57e9b-104">Funções personalizadas estendem os recursos do Excel usando um novo tempo de execução do JavaScript que usa um mecanismo de JavaScript em área restrita em vez de um navegador da web.</span><span class="sxs-lookup"><span data-stu-id="57e9b-104">Custom functions extend Excel’s capabilities by using a new JavaScript runtime that uses a sandboxed JavaScript engine rather than a web browser.</span></span> <span data-ttu-id="57e9b-105">Como as funções personalizadas não precisam renderizar elementos de interface do usuário, o novo tempo de execução do JavaScript é otimizado para fazer cálculos, permitindo que você execute milhares de funções personalizadas simultaneamente.</span><span class="sxs-lookup"><span data-stu-id="57e9b-105">Because custom functions do not need to render UI elements, the new JavaScript runtime is optimized for performing calculations, enabling you to run thousands of custom functions simultaneously.</span></span>

## <a name="key-facts-about-the-new-javascript-runtime"></a><span data-ttu-id="57e9b-106">Fatos importantes sobre o novo tempo de execução do JavaScript</span><span class="sxs-lookup"><span data-stu-id="57e9b-106">Key facts about the new JavaScript runtime</span></span> 

<span data-ttu-id="57e9b-107">Somente as funções personalizadas de um suplemento usam o novo tempo de execução do JavaScript descrito neste artigo.</span><span class="sxs-lookup"><span data-stu-id="57e9b-107">Only custom functions within an add-in will use the new JavaScript runtime that's described in this article.</span></span> <span data-ttu-id="57e9b-108">Se um suplemento incluir outros componentes, como painéis de tarefas e outros elementos de interface do usuário, além das funções personalizadas, esses outros componentes do suplemento continuarão operando no tempo de execução de exibição da Web com aparência de navegador.</span><span class="sxs-lookup"><span data-stu-id="57e9b-108">If an add-in includes other components such as task panes and other UI elements, in addition to custom functions, these other components of the add-in will continue to run in the browser-like WebView runtime.</span></span>  <span data-ttu-id="57e9b-109">Além disso:</span><span class="sxs-lookup"><span data-stu-id="57e9b-109">Additionally:</span></span> 

- <span data-ttu-id="57e9b-110">O tempo de execução do JavaScript não fornece acesso ao Document Object Model (DOM) ou a bibliotecas de suporte, como jQuery, que dependem do DOM.</span><span class="sxs-lookup"><span data-stu-id="57e9b-110">The JavaScript runtime does not provide access to the Document Object Model (DOM) or support libraries like jQuery that rely on the DOM.</span></span>

- <span data-ttu-id="57e9b-111">Uma função personalizada que é definida em um arquivo JavaScript de um suplemento pode retornar um `Promise` regular do JavaScript em vez de retornar `OfficeExtension.Promise`.</span><span class="sxs-lookup"><span data-stu-id="57e9b-111">A custom function that's defined in an add-in's JavaScript file can return a regular JavaScript `Promise` instead of returning `OfficeExtension.Promise`.</span></span>  

- <span data-ttu-id="57e9b-112">O arquivo JSON que especifica a função personalizada metatdata não precisa especificar **sync** ou **async** nas **opções**.</span><span class="sxs-lookup"><span data-stu-id="57e9b-112">The JSON file that specifies custom function metatdata does not need to specify **sync** or **async** within **options**.</span></span>

## <a name="new-apis"></a><span data-ttu-id="57e9b-113">Novas APIs</span><span class="sxs-lookup"><span data-stu-id="57e9b-113">New Excel JavaScript APIs</span></span> 

<span data-ttu-id="57e9b-114">O tempo de execução do JavaScript que é usado pelas funções personalizadas tem as seguintes APIs:</span><span class="sxs-lookup"><span data-stu-id="57e9b-114">The JavaScript runtime that's used by custom functions has the following APIs:</span></span>

- [<span data-ttu-id="57e9b-115">XHR</span><span class="sxs-lookup"><span data-stu-id="57e9b-115">XHR</span></span>](#xhr)
- [<span data-ttu-id="57e9b-116">WebSockets</span><span class="sxs-lookup"><span data-stu-id="57e9b-116">WebSockets</span></span>](#websockets)
- [<span data-ttu-id="57e9b-117">AsyncStorage</span><span class="sxs-lookup"><span data-stu-id="57e9b-117">AsyncStorage</span></span>](#asyncstorage)
- [<span data-ttu-id="57e9b-118">API de diálogo</span><span class="sxs-lookup"><span data-stu-id="57e9b-118">Dialog API requirement sets</span></span>](#dialog-api)

### <a name="xhr"></a><span data-ttu-id="57e9b-119">XHR</span><span class="sxs-lookup"><span data-stu-id="57e9b-119">XHR</span></span>

<span data-ttu-id="57e9b-120">XHR significa [XmlHttpRequest](https://developer.mozilla.org/en-US/docs/Web/API/XMLHttpRequest), uma API da Web padrão que emite solicitações HTTP para interagir com os servidores.</span><span class="sxs-lookup"><span data-stu-id="57e9b-120">XHR stands for [XmlHttpRequest](https://developer.mozilla.org/en-US/docs/Web/API/XMLHttpRequest), a standard web API that issues HTTP requests to interact with servers.</span></span> <span data-ttu-id="57e9b-121">No novo tempo de execução do JavaScript, XHR implementa medidas adicionais de segurança, exigindo a [Mesma diretiva de origem](https://developer.mozilla.org/en-US/docs/Web/Security/Same-origin_policy) e [CORS](https://www.w3.org/TR/cors/) simples.</span><span class="sxs-lookup"><span data-stu-id="57e9b-121">In the new JavaScript runtime, XHR implements additional security measures by requiring [Same Origin Policy](https://developer.mozilla.org/en-US/docs/Web/Security/Same-origin_policy) and simple [CORS](https://www.w3.org/TR/cors/).</span></span>  

<span data-ttu-id="57e9b-122">No exemplo de código a seguir, a função `getTemperature()` envia uma solicitação da web para obter a temperatura de uma determinada área com base na ID de termômetro.</span><span class="sxs-lookup"><span data-stu-id="57e9b-122">In the following code sample, the `getTemperature()` function sends a web request to get the temperature of a particular area based on thermometer ID.</span></span> <span data-ttu-id="57e9b-123">A função `sendWebRequest()` usa XHR para emitir uma solicitação `GET` para um ponto de extremidade que pode fornecer os dados.</span><span class="sxs-lookup"><span data-stu-id="57e9b-123">The `sendWebRequest()` function uses XHR to issue a `GET` request to an endpoint that can provide the data.</span></span>  

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

### <a name="websockets"></a><span data-ttu-id="57e9b-124">WebSockets</span><span class="sxs-lookup"><span data-stu-id="57e9b-124">WebSockets</span></span>

<span data-ttu-id="57e9b-125">[WebSockets](https://developer.mozilla.org/en-US/docs/Web/API/WebSockets_API) é um protocolo de rede que cria a comunicação em tempo real entre um servidor e um ou mais clientes.</span><span class="sxs-lookup"><span data-stu-id="57e9b-125">[WebSockets](https://developer.mozilla.org/en-US/docs/Web/API/WebSockets_API) is a networking protocol that creates real-time communication between a server and one or more clients.</span></span> <span data-ttu-id="57e9b-126">Ele é frequentemente usado para aplicativos de bate-papo porque permite que você possa ler e gravar texto simultaneamente.</span><span class="sxs-lookup"><span data-stu-id="57e9b-126">It is often used for chat applications because it allows you to read and write text simultaneously.</span></span>  

<span data-ttu-id="57e9b-127">Como mostra o exemplo de código a seguir, as funções personalizadas podem usar WebSockets.</span><span class="sxs-lookup"><span data-stu-id="57e9b-127">As shown in the following code sample, custom functions can use WebSockets.</span></span> <span data-ttu-id="57e9b-128">Neste exemplo, o WebSocket registra cada mensagem que recebe.</span><span class="sxs-lookup"><span data-stu-id="57e9b-128">In this example, the WebSocket logs each message that it receives.</span></span>

```ts
const ws = new WebSocket('wss://bundles.office.com');
ws.onmessage = (message) => {
    console.log(`Received: ${message}`);
};
ws.onerror = (error) => {
    console.err(`Failed: ${error}`);
}
```

### <a name="asyncstorage"></a><span data-ttu-id="57e9b-129">AsyncStorage</span><span class="sxs-lookup"><span data-stu-id="57e9b-129">AsyncStorage</span></span>

<span data-ttu-id="57e9b-130">AsyncStorage é um sistema de armazenamento de chave-valor que pode ser usado para armazenar os tokens de autenticação.</span><span class="sxs-lookup"><span data-stu-id="57e9b-130">AsyncStorage is a key-value storage system that can be used to store authentication tokens.</span></span> <span data-ttu-id="57e9b-131">Ele é:</span><span class="sxs-lookup"><span data-stu-id="57e9b-131">It is framework-agnostic.</span></span>

- <span data-ttu-id="57e9b-132">Persistente</span><span class="sxs-lookup"><span data-stu-id="57e9b-132">Persistent Property</span></span>
- <span data-ttu-id="57e9b-133">Não encriptado</span><span class="sxs-lookup"><span data-stu-id="57e9b-133">Unencrypted</span></span>
- <span data-ttu-id="57e9b-134">Assíncrono</span><span class="sxs-lookup"><span data-stu-id="57e9b-134">Asynchronous calls</span></span>

<span data-ttu-id="57e9b-135">AsyncStorage fica disponível globalmente para todas as partes do seu suplemento.</span><span class="sxs-lookup"><span data-stu-id="57e9b-135">AsyncStorage is globally available to all parts of your add-in.</span></span> <span data-ttu-id="57e9b-136">Para funções personalizadas, `AsyncStorage` é exposto como um objeto global.</span><span class="sxs-lookup"><span data-stu-id="57e9b-136">For custom functions, `AsyncStorage` is exposed as a global object.</span></span> <span data-ttu-id="57e9b-137">(Para outras partes do seu suplemento, como painéis de tarefas e outros elementos que usam o tempo de execução de exibição da Web, AsyncStorage é exposto por meio do `OfficeRuntime`.) Cada suplemento tem sua própria partição de armazenamento, com um tamanho padrão de 5 MB.</span><span class="sxs-lookup"><span data-stu-id="57e9b-137">(For other parts of your add-in, such as task panes and other elements that use the WebView runtime, AsyncStorage is exposed through `OfficeRuntime`.) Each add-in has its own storage partition, with a default size of 5MB.</span></span> 

<span data-ttu-id="57e9b-138">Os métodos a seguir estão disponíveis no objeto `AsyncStorage`:</span><span class="sxs-lookup"><span data-stu-id="57e9b-138">The following methods are available on the `AsyncStorage` object:</span></span>
 
 - `getItem`
 - `setItem`
 - `removeItem`
 - `clear`
 - `getAllKeys`
 - `flushGetRequests`
 - `multiGet`
 - `multiSet`
 - `multiRemove`
 
<span data-ttu-id="57e9b-139">Neste momento, os métodos `mergeItem` e `multiMerge` não são suportados.</span><span class="sxs-lookup"><span data-stu-id="57e9b-139">At this time, the `mergeItem` and `multiMerge` methods are not supported.</span></span>

<span data-ttu-id="57e9b-140">O seguinte código de amostra chama a função `AsyncStorage.getItem` para recuperar um valor de armazenamento.</span><span class="sxs-lookup"><span data-stu-id="57e9b-140">The following code sample calls the `AsyncStorage.getItem` function to retrieve a value from storage.</span></span>

```js
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

### <a name="dialog-api"></a><span data-ttu-id="57e9b-141">API de diálogo</span><span class="sxs-lookup"><span data-stu-id="57e9b-141">Dialog API scenarios</span></span>

<span data-ttu-id="57e9b-142">A API de diálogo permite que você abra uma caixa de diálogo que solicita a entrada do usuário.</span><span class="sxs-lookup"><span data-stu-id="57e9b-142">The Dialog API enables you to open a dialog box that prompts user sign-in.</span></span> <span data-ttu-id="57e9b-143">Você pode usar a API de diálogo para exigir a autenticação de usuário por meio de um recurso externo, como Google ou Facebook, para que o usuário possa usar sua função.</span><span class="sxs-lookup"><span data-stu-id="57e9b-143">You can use the Dialog API to require user authentication through an outside resource, such as Google or Facebook, before the user can use your function.</span></span>   

<span data-ttu-id="57e9b-144">No exemplo de código a seguir, o método `getTokenViaDialog()` usa o método `displayWebDialog()` da API de diálogo para abrir uma caixa de diálogo.</span><span class="sxs-lookup"><span data-stu-id="57e9b-144">In the following code sample, the `getTokenViaDialog()` method uses the Dialog API’s `displayWebDialog()` method to open a dialog box.</span></span>

```js
// Get auth token before calling my service, a hypothetical API that will deliver a stock price based on stock ticker string, such as "MSFT"
 
function getStock (ticker) {
  return new Promise(function (resolve, reject) {
    // Get a token
    getToken("https://myauthurl")
    .then(function (token) {
      
      // Use token to get stock price
      fetch("https://myservice.com/?token=token&ticker= + ticker")
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
> <span data-ttu-id="57e9b-145">A API de diálogo descrita nesta seção faz parte do novo tempo de execução do JavaScript para funções personalizadas e pode ser usada somente nas funções personalizadas.</span><span class="sxs-lookup"><span data-stu-id="57e9b-145">The Dialog API described in this section is part of the new JavaScript runtime for custom functions and can be used only within custom functions.</span></span> <span data-ttu-id="57e9b-146">Essa API é diferente da [API de diálogo](../develop/dialog-api-in-office-add-ins.md) que pode ser usada nos painéis de tarefas e comandos do suplemento.</span><span class="sxs-lookup"><span data-stu-id="57e9b-146">This API is different from the [Dialog API](../develop/dialog-api-in-office-add-ins.md) that can be used within task panes and add-in commands.</span></span>

## <a name="see-also"></a><span data-ttu-id="57e9b-147">Confira também</span><span class="sxs-lookup"><span data-stu-id="57e9b-147">See also</span></span>

* [<span data-ttu-id="57e9b-148">Criar funções personalizadas no Excel</span><span class="sxs-lookup"><span data-stu-id="57e9b-148">Create custom functions in Excel (Preview)</span></span>](custom-functions-overview.md)
* [<span data-ttu-id="57e9b-149">Metadados de funções personalizadas</span><span class="sxs-lookup"><span data-stu-id="57e9b-149">Custom functions metadata</span></span>](custom-functions-json.md)
* [<span data-ttu-id="57e9b-150">Melhores práticas de funções personalizadas</span><span class="sxs-lookup"><span data-stu-id="57e9b-150">Custom functions best practices</span></span>](custom-functions-best-practices.md)