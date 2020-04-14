---
ms.date: 04/13/2020
description: Entenda os principais cenários de desenvolvimento de funções personalizadas do Excel que usam o novo tempo de execução do JavaScript.
title: Tempo de execução de funções personalizadas do Excel
localization_priority: Normal
ms.openlocfilehash: dc049aa681ae4f7664d5bd92f925e7566c0d7103
ms.sourcegitcommit: 118e8bcbcfb73c93e2053bda67fe8dd20799b170
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 04/13/2020
ms.locfileid: "43241039"
---
# <a name="runtime-for-excel-custom-functions"></a><span data-ttu-id="4eb45-103">Tempo de execução de funções personalizadas do Excel</span><span class="sxs-lookup"><span data-stu-id="4eb45-103">Runtime for Excel custom functions</span></span>

<span data-ttu-id="4eb45-104">Funções personalizadas usam um novo tempo de execução do JavaScript, diferente do tempo de execução usado por outras partes de um suplemento, como o painel de tarefas ou outros elementos da interface do usuário.</span><span class="sxs-lookup"><span data-stu-id="4eb45-104">Custom functions use a new JavaScript runtime that differs from the runtime used by other parts of an add-in, such as the task pane or other UI elements.</span></span> <span data-ttu-id="4eb45-105">Esse tempo de execução do JavaScript foi projetado para otimizar o desempenho de cálculos em funções personalizadas, e expõe as novas APIs disponíveis para executar ações comuns baseadas na Web, dentro de funções personalizadas, como solicitação de dados externos ou troca de dados por meio de uma conexão persistente com um servidor.</span><span class="sxs-lookup"><span data-stu-id="4eb45-105">This JavaScript runtime is designed to optimize performance of calculations in custom functions and exposes new APIs that you can use to perform common web-based actions within custom functions such as requesting external data or exchanging data over a persistent connection with a server.</span></span>

[!include[Excel custom functions note](../includes/excel-custom-functions-note.md)]

<span data-ttu-id="4eb45-106">O tempo de execução do JavaScript também fornece acesso às novas APIs no namespace `OfficeRuntime` que pode ser usado em funções personalizadas ou por outras partes de um suplemento para armazenar dados ou exibir uma caixa de diálogo.</span><span class="sxs-lookup"><span data-stu-id="4eb45-106">The JavaScript runtime also provides access to new APIs in the `OfficeRuntime` namespace that can be used within custom functions or by other parts of an add-in to store data or display a dialog box.</span></span> <span data-ttu-id="4eb45-107">Este artigo mostra como usar essas APIs em funções personalizadas e descreve considerações adicionais para o desenvolvimento de funções personalizadas.</span><span class="sxs-lookup"><span data-stu-id="4eb45-107">This article describes how to use these APIs within custom functions and also outlines additional considerations to keep in mind as you develop custom functions.</span></span>

## <a name="requesting-external-data"></a><span data-ttu-id="4eb45-108">Como solicitar dados externos</span><span class="sxs-lookup"><span data-stu-id="4eb45-108">Requesting external data</span></span>

<span data-ttu-id="4eb45-109">É possível solicitar dados externos em uma função personalizada por meio de uma API, como a API [Fetch](https://developer.mozilla.org/en-US/docs/Web/API/Fetch_API), ou por meio de um objeto [XmlHttpRequest (XHR)](https://developer.mozilla.org/en-US/docs/Web/API/XMLHttpRequest), uma API Web padrão que envia solicitações HTTP para interagir com os servidores.</span><span class="sxs-lookup"><span data-stu-id="4eb45-109">Within a custom function, you can request external data by using an API like [Fetch](https://developer.mozilla.org/en-US/docs/Web/API/Fetch_API) or by using [XmlHttpRequest (XHR)](https://developer.mozilla.org/en-US/docs/Web/API/XMLHttpRequest), a standard web API that issues HTTP requests to interact with servers.</span></span>

<span data-ttu-id="4eb45-110">Dentro do tempo de execução do JavaScript usado por funções personalizadas, o XHR implementa medidas de segurança adicionais exigindo a [mesma política de origem](https://developer.mozilla.org/en-US/docs/Web/Security/Same-origin_policy) e o [CORS](https://www.w3.org/TR/cors/)simples.</span><span class="sxs-lookup"><span data-stu-id="4eb45-110">Within the JavaScript runtime used by custom functions, XHR implements additional security measures by requiring [Same Origin Policy](https://developer.mozilla.org/en-US/docs/Web/Security/Same-origin_policy) and simple [CORS](https://www.w3.org/TR/cors/).</span></span>

<span data-ttu-id="4eb45-111">Observe que uma implementação CORS simples não pode usar cookies e é compatível apenas com métodos simples (GET, HEAD, POST).</span><span class="sxs-lookup"><span data-stu-id="4eb45-111">Note that a simple CORS implementation cannot use cookies and only supports simple methods (GET, HEAD, POST).</span></span> <span data-ttu-id="4eb45-112">A CORS simples aceita cabeçalhos simples com nomes de campos `Accept`, `Accept-Language`, `Content-Language`.</span><span class="sxs-lookup"><span data-stu-id="4eb45-112">Simple CORS accepts simple headers with field names `Accept`, `Accept-Language`, `Content-Language`.</span></span> <span data-ttu-id="4eb45-113">Você também pode usar um `Content-Type` cabeçalho no CORS simples, desde que o tipo de conteúdo `application/x-www-form-urlencoded`seja `text/plain`, ou `multipart/form-data`.</span><span class="sxs-lookup"><span data-stu-id="4eb45-113">You can also use a `Content-Type` header in simple CORS, provided that the content type is `application/x-www-form-urlencoded`, `text/plain`, or `multipart/form-data`.</span></span>

### <a name="xhr-example"></a><span data-ttu-id="4eb45-114">Exemplo de XHR</span><span class="sxs-lookup"><span data-stu-id="4eb45-114">XHR example</span></span>

<span data-ttu-id="4eb45-115">No código de exemplo a seguir, a função `getTemperature` chama a função `sendWebRequest` para obter a temperatura de uma área específica, de acordo com a ID do termômetro.</span><span class="sxs-lookup"><span data-stu-id="4eb45-115">In the following code sample, the `getTemperature` function calls the `sendWebRequest` function to get the temperature of a particular area based on thermometer ID.</span></span> <span data-ttu-id="4eb45-116">A função `sendWebRequest` usa XHR para emitir uma solicitação `GET` para um ponto de extremidade que fornece os dados.</span><span class="sxs-lookup"><span data-stu-id="4eb45-116">The `sendWebRequest` function uses XHR to issue a `GET` request to an endpoint that can provide the data.</span></span>

> [!NOTE] 
> <span data-ttu-id="4eb45-117">Se usar fetch ou XHR, uma nova `Promise` JavaScript será retornada.</span><span class="sxs-lookup"><span data-stu-id="4eb45-117">When using fetch or XHR, a new JavaScript `Promise` is returned.</span></span> <span data-ttu-id="4eb45-118">Antes de setembro de 2018, era necessário especificar `OfficeExtension.Promise` para usar promessas na API JavaScript para Office, mas agora, basta usar um `Promise` JavaScript.</span><span class="sxs-lookup"><span data-stu-id="4eb45-118">Prior to September 2018, you had to specify `OfficeExtension.Promise` to use promises within the Office JavaScript API, but now you can simply use a JavaScript `Promise`.</span></span>

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
        
        //set Content-Type to application/text. Application/json is not currently supported with Simple CORS
        xhttp.setRequestHeader("Content-Type", "application/text");
        xhttp.open("GET", "https://contoso.com/temperature/" + thermometerID), true)
        xhttp.send();  
    }
}
```

## <a name="receiving-data-via-websockets"></a><span data-ttu-id="4eb45-119">Como receber dados por meio de WebSockets</span><span class="sxs-lookup"><span data-stu-id="4eb45-119">Receiving data via WebSockets</span></span>

<span data-ttu-id="4eb45-120">Em uma função personalizada, é possível usar [WebSockets](https://developer.mozilla.org/en-US/docs/Web/API/WebSockets_API) para trocar dados por meio de uma conexão persistente com um servidor.</span><span class="sxs-lookup"><span data-stu-id="4eb45-120">Within a custom function, you can use [WebSockets](https://developer.mozilla.org/en-US/docs/Web/API/WebSockets_API) to exchange data over a persistent connection with a server.</span></span> <span data-ttu-id="4eb45-121">Usando WebSockets, a função personalizada pode abrir uma conexão com um servidor e, em seguida, receber mensagens do servidor automaticamente, quando determinados eventos ocorrerem, sem precisar consultar explicitamente os dados do servidor.</span><span class="sxs-lookup"><span data-stu-id="4eb45-121">By using WebSockets, your custom function can open a connection with a server and then automatically receive messages from the server when certain events occur, without having to explicitly poll the server for data.</span></span>

### <a name="websockets-example"></a><span data-ttu-id="4eb45-122">Exemplo de WebSockets</span><span class="sxs-lookup"><span data-stu-id="4eb45-122">WebSockets example</span></span>

<span data-ttu-id="4eb45-123">O código de exemplo a seguir estabelece uma conexão `WebSocket` e registra cada mensagem de entrada do servidor.</span><span class="sxs-lookup"><span data-stu-id="4eb45-123">The following code sample establishes a `WebSocket` connection and then logs each incoming message from the server.</span></span>

```js
const ws = new WebSocket('wss://bundles.office.com');
ws.onmessage = function (message) {
    console.log(`Received: ${message}`);
}
ws.onerror = function (error) {
    console.err(`Failed: ${error}`);
}
```

## <a name="storing-and-accessing-data"></a><span data-ttu-id="4eb45-124">Como armazenar e acessar os dados</span><span class="sxs-lookup"><span data-stu-id="4eb45-124">Storing and accessing data</span></span>

<span data-ttu-id="4eb45-125">Em uma função personalizada (ou em outras partes de um suplemento), você pode armazenar e acessar dados usando o objeto `OfficeRuntime.storage`.</span><span class="sxs-lookup"><span data-stu-id="4eb45-125">Within a custom function (or within any other part of an add-in), you can store and access data by using the `OfficeRuntime.storage` object.</span></span> <span data-ttu-id="4eb45-126">`Storage` é um sistema de armazenamento de chave-valor persistente e não criptografado, que fornece uma alternativa para [localStorage](https://developer.mozilla.org/en-US/docs/Web/API/Window/localStorage), que não pode ser usado em funções personalizadas.</span><span class="sxs-lookup"><span data-stu-id="4eb45-126">`Storage` is a persistent, unencrypted, key-value storage system that provides an alternative to [localStorage](https://developer.mozilla.org/en-US/docs/Web/API/Window/localStorage), which cannot be used within custom functions.</span></span> <span data-ttu-id="4eb45-127">`Storage`o oferece 10 MB de dados por domínio.</span><span class="sxs-lookup"><span data-stu-id="4eb45-127">`Storage` offers 10 MB of data per domain.</span></span> <span data-ttu-id="4eb45-128">Os domínios podem ser compartilhados por mais de um suplemento.</span><span class="sxs-lookup"><span data-stu-id="4eb45-128">Domains can be shared by more than one add-in.</span></span>

<span data-ttu-id="4eb45-129">`Storage` é uma solução de armazenamento compartilhado, o que significa que várias partes de um suplemento podem acessar os mesmos dados.</span><span class="sxs-lookup"><span data-stu-id="4eb45-129">`Storage` is intended as a shared storage solution, meaning multiple parts of an add-in are able to access the same data.</span></span> <span data-ttu-id="4eb45-130">Por exemplo, tokens para autenticação de usuário podem ser armazenados em `storage`, já que ele pode ser acessado tanto por uma função personalizada quanto por elementos da interface do usuário de um suplemento, como um painel de tarefas.</span><span class="sxs-lookup"><span data-stu-id="4eb45-130">For example, tokens for user authentication may be stored in `storage` because it can be accessed by both a custom function and add-in UI elements such as a task pane.</span></span> <span data-ttu-id="4eb45-131">Da mesma forma, se dois suplementos compartilham o mesmo domínio (por exemplo, `www.contoso.com/addin1` `www.contoso.com/addin2`), eles também podem compartilhar informações de frente e para trás `storage`.</span><span class="sxs-lookup"><span data-stu-id="4eb45-131">Similarly, if two add-ins share the same domain (for example, `www.contoso.com/addin1`, `www.contoso.com/addin2`), they are also permitted to share information back and forth through `storage`.</span></span> <span data-ttu-id="4eb45-132">Observe que os suplementos que possuem subdomínios diferentes terão instâncias diferentes `storage` (por exemplo, `subdomain.contoso.com/addin1`, `differentsubdomain.contoso.com/addin2`).</span><span class="sxs-lookup"><span data-stu-id="4eb45-132">Note that add-ins which have different subdomains will have different instances of `storage` (for example, `subdomain.contoso.com/addin1`, `differentsubdomain.contoso.com/addin2`).</span></span>

<span data-ttu-id="4eb45-133">Como `storage` pode ser um local compartilhado, é importante notar que é possível substituir os pares chave-valor.</span><span class="sxs-lookup"><span data-stu-id="4eb45-133">Because `storage` can be a shared location, it is important to realize that it is possible to override key-value pairs.</span></span>

<span data-ttu-id="4eb45-134">Os métodos a seguir estão disponíveis no objeto `storage`:</span><span class="sxs-lookup"><span data-stu-id="4eb45-134">The following methods are available on the `storage` object:</span></span>

 - `getItem`
 - `getItems`
 - `setItem`
 - `setItems`
 - `removeItem`
 - `removeItems`
 - `getKeys`

<span data-ttu-id="4eb45-135">.</span><span class="sxs-lookup"><span data-stu-id="4eb45-135">.</span></span>[!NOTE]
> <span data-ttu-id="4eb45-136">Não há nenhum método para limpar todas as informações (como `clear`).</span><span class="sxs-lookup"><span data-stu-id="4eb45-136">There's no method for clearing all information (such as `clear`).</span></span> <span data-ttu-id="4eb45-137">Em vez disso, use `removeItems` para remover várias entradas de uma só vez.</span><span class="sxs-lookup"><span data-stu-id="4eb45-137">Instead, you should instead use `removeItems` to remove multiple entries at a time.</span></span>

### <a name="officeruntimestorage-example"></a><span data-ttu-id="4eb45-138">Exemplo de OfficeRuntime. Storage</span><span class="sxs-lookup"><span data-stu-id="4eb45-138">OfficeRuntime.storage example</span></span>

<span data-ttu-id="4eb45-139">O exemplo de código a seguir `OfficeRuntime.storage.setItem` chama a função para definir uma chave e `storage`um valor para.</span><span class="sxs-lookup"><span data-stu-id="4eb45-139">The following code sample calls the `OfficeRuntime.storage.setItem` function to set a key and value into `storage`.</span></span>

```js
function StoreValue(key, value) {

  return OfficeRuntime.storage.setItem(key, value).then(function (result) {
      return "Success: Item with key '" + key + "' saved to storage.";
  }, function (error) {
      return "Error: Unable to save item with key '" + key + "' to storage. " + error;
  });
}
```

## <a name="additional-considerations"></a><span data-ttu-id="4eb45-140">Considerações adicionais</span><span class="sxs-lookup"><span data-stu-id="4eb45-140">Additional considerations</span></span>

<span data-ttu-id="4eb45-141">Para criar um suplemento que será executado em várias plataformas (um dos principais locatários de Suplementos do Office), você não deve acessar o DOM (Modelo de Objeto do Documento) em funções personalizadas nem usar bibliotecas, como a jQuery, que dependem do DOM.</span><span class="sxs-lookup"><span data-stu-id="4eb45-141">In order to create an add-in that will run on multiple platforms (one of the key tenants of Office Add-ins), you should not access the Document Object Model (DOM) in custom functions or use libraries like jQuery that rely on the DOM.</span></span> <span data-ttu-id="4eb45-142">No Excel no Windows, onde as funções personalizadas usam o tempo de execução do JavaScript, as funções personalizadas não podem acessar o DOM.</span><span class="sxs-lookup"><span data-stu-id="4eb45-142">In Excel on Windows, where custom functions use the JavaScript runtime, custom functions cannot access the DOM.</span></span>

## <a name="next-steps"></a><span data-ttu-id="4eb45-143">Próximas etapas</span><span class="sxs-lookup"><span data-stu-id="4eb45-143">Next steps</span></span>
<span data-ttu-id="4eb45-144">Saiba como [realizar solicitações da Web com funções personalizadas](custom-functions-web-reqs.md).</span><span class="sxs-lookup"><span data-stu-id="4eb45-144">Learn how to [perform web requests with custom functions](custom-functions-web-reqs.md).</span></span>

## <a name="see-also"></a><span data-ttu-id="4eb45-145">Confira também</span><span class="sxs-lookup"><span data-stu-id="4eb45-145">See also</span></span>

* [<span data-ttu-id="4eb45-146">Criar funções personalizadas no Excel</span><span class="sxs-lookup"><span data-stu-id="4eb45-146">Create custom functions in Excel</span></span>](custom-functions-overview.md)
* [<span data-ttu-id="4eb45-147">Arquitetura de funções personalizadas</span><span class="sxs-lookup"><span data-stu-id="4eb45-147">Custom functions architecture</span></span>](custom-functions-architecture.md)
* [<span data-ttu-id="4eb45-148">Exibir uma caixa de diálogo em funções personalizadas</span><span class="sxs-lookup"><span data-stu-id="4eb45-148">Display a dialog in custom functions</span></span>](custom-functions-dialog.md)
* [<span data-ttu-id="4eb45-149">Tutorial de funções personalizadas</span><span class="sxs-lookup"><span data-stu-id="4eb45-149">Custom functions tutorial</span></span>](../tutorials/excel-tutorial-create-custom-functions.md)
