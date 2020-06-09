---
ms.date: 05/17/2020
description: Entenda as funções personalizadas do Excel que não usam um painel de tarefas e seu tempo de execução JavaScript específico.
title: Tempo de execução para funções personalizadas do Excel sem interface do usuário
localization_priority: Normal
ms.openlocfilehash: 5cb9aa480d6923d31434d58a9683e9a9f5d48458
ms.sourcegitcommit: be23b68eb661015508797333915b44381dd29bdb
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 06/08/2020
ms.locfileid: "44609640"
---
# <a name="runtime-for-ui-less-excel-custom-functions"></a><span data-ttu-id="fc85f-103">Tempo de execução para funções personalizadas do Excel sem interface do usuário</span><span class="sxs-lookup"><span data-stu-id="fc85f-103">Runtime for UI-less Excel custom functions</span></span>

<span data-ttu-id="fc85f-104">As funções personalizadas que não usam um painel de tarefas (funções personalizadas sem interface do usuário) usam um tempo de execução do JavaScript projetado para otimizar o desempenho dos cálculos.</span><span class="sxs-lookup"><span data-stu-id="fc85f-104">Custom functions that don't use a task pane (UI-less custom functions) use a JavaScript runtime that is designed to optimize performance of calculations.</span></span>

[!include[Excel custom functions note](../includes/excel-custom-functions-note.md)]

[!include[Shared runtime note](../includes/shared-runtime-note.md)]

<span data-ttu-id="fc85f-105">Este tempo de execução JavaScript fornece acesso a APIs no `OfficeRuntime` namespace que podem ser usadas por funções personalizadas sem interface do usuário e o painel de tarefas para armazenar dados.</span><span class="sxs-lookup"><span data-stu-id="fc85f-105">This JavaScript runtime provides access to APIs in the `OfficeRuntime` namespace that can be used by UI-less custom functions and the task pane to store data.</span></span>

## <a name="requesting-external-data"></a><span data-ttu-id="fc85f-106">Como solicitar dados externos</span><span class="sxs-lookup"><span data-stu-id="fc85f-106">Requesting external data</span></span>

<span data-ttu-id="fc85f-107">Dentro de uma função personalizada sem interface do usuário, você pode solicitar dados externos usando uma API como [Fetch](https://developer.mozilla.org/en-US/docs/Web/API/Fetch_API) ou usando [XMLHttpRequest (XHR)](https://developer.mozilla.org/en-US/docs/Web/API/XMLHttpRequest), uma API Web padrão que emite solicitações HTTP para interagir com os servidores.</span><span class="sxs-lookup"><span data-stu-id="fc85f-107">Within a UI-less custom function, you can request external data by using an API like [Fetch](https://developer.mozilla.org/en-US/docs/Web/API/Fetch_API) or by using [XmlHttpRequest (XHR)](https://developer.mozilla.org/en-US/docs/Web/API/XMLHttpRequest), a standard web API that issues HTTP requests to interact with servers.</span></span>

<span data-ttu-id="fc85f-108">Esteja ciente de que as funções sem interface do usuário devem usar medidas de segurança adicionais ao fazer XMLHttpRequests, exigindo a [mesma política de origem](https://developer.mozilla.org/en-US/docs/Web/Security/Same-origin_policy) e [CORS](https://www.w3.org/TR/cors/)simples.</span><span class="sxs-lookup"><span data-stu-id="fc85f-108">Be aware that UI-less functions must use additional security measures when making XmlHttpRequests, requiring [Same Origin Policy](https://developer.mozilla.org/en-US/docs/Web/Security/Same-origin_policy) and simple [CORS](https://www.w3.org/TR/cors/).</span></span>

<span data-ttu-id="fc85f-109">Uma implementação CORS simples não pode usar cookies e só oferece suporte a métodos simples (GET, HEAD, POST).</span><span class="sxs-lookup"><span data-stu-id="fc85f-109">A simple CORS implementation cannot use cookies and only supports simple methods (GET, HEAD, POST).</span></span> <span data-ttu-id="fc85f-110">A CORS simples aceita cabeçalhos simples com nomes de campos `Accept`, `Accept-Language`, `Content-Language`.</span><span class="sxs-lookup"><span data-stu-id="fc85f-110">Simple CORS accepts simple headers with field names `Accept`, `Accept-Language`, `Content-Language`.</span></span> <span data-ttu-id="fc85f-111">Você também pode usar um `Content-Type` cabeçalho no CORS simples, desde que o tipo de conteúdo seja `application/x-www-form-urlencoded` , `text/plain` ou `multipart/form-data` .</span><span class="sxs-lookup"><span data-stu-id="fc85f-111">You can also use a `Content-Type` header in simple CORS, provided that the content type is `application/x-www-form-urlencoded`, `text/plain`, or `multipart/form-data`.</span></span>

## <a name="storing-and-accessing-data"></a><span data-ttu-id="fc85f-112">Como armazenar e acessar os dados</span><span class="sxs-lookup"><span data-stu-id="fc85f-112">Storing and accessing data</span></span>

<span data-ttu-id="fc85f-113">Dentro de uma função personalizada sem interface do usuário, você pode armazenar e acessar dados usando o `OfficeRuntime.storage` objeto.</span><span class="sxs-lookup"><span data-stu-id="fc85f-113">Within a UI-less custom function, you can store and access data by using the `OfficeRuntime.storage` object.</span></span> <span data-ttu-id="fc85f-114">`Storage`é um sistema de armazenamento de valor chave persistente, não criptografado que fornece uma alternativa para o [localStorage](https://developer.mozilla.org/en-US/docs/Web/API/Window/localStorage), que não pode ser usado por funções personalizadas sem interface do usuário.</span><span class="sxs-lookup"><span data-stu-id="fc85f-114">`Storage` is a persistent, unencrypted, key-value storage system that provides an alternative to [localStorage](https://developer.mozilla.org/en-US/docs/Web/API/Window/localStorage), which cannot be used by UI-less custom functions.</span></span> <span data-ttu-id="fc85f-115">`Storage`o oferece 10 MB de dados por domínio.</span><span class="sxs-lookup"><span data-stu-id="fc85f-115">`Storage` offers 10 MB of data per domain.</span></span> <span data-ttu-id="fc85f-116">Os domínios podem ser compartilhados por mais de um suplemento.</span><span class="sxs-lookup"><span data-stu-id="fc85f-116">Domains can be shared by more than one add-in.</span></span>

<span data-ttu-id="fc85f-117">`Storage` é uma solução de armazenamento compartilhado, o que significa que várias partes de um suplemento podem acessar os mesmos dados.</span><span class="sxs-lookup"><span data-stu-id="fc85f-117">`Storage` is intended as a shared storage solution, meaning multiple parts of an add-in are able to access the same data.</span></span> <span data-ttu-id="fc85f-118">Por exemplo, os tokens para autenticação de usuário podem ser armazenados em `storage` porque podem ser acessados por uma função personalizada sem interface e elementos de interface do usuário de suplemento, como um painel de tarefas.</span><span class="sxs-lookup"><span data-stu-id="fc85f-118">For example, tokens for user authentication may be stored in `storage` because it can be accessed by both a UI-less custom function and add-in UI elements such as a task pane.</span></span> <span data-ttu-id="fc85f-119">Da mesma forma, se dois suplementos compartilham o mesmo domínio (por exemplo, `www.contoso.com/addin1` `www.contoso.com/addin2` ), eles também podem compartilhar informações de frente e para trás `storage` .</span><span class="sxs-lookup"><span data-stu-id="fc85f-119">Similarly, if two add-ins share the same domain (for example, `www.contoso.com/addin1`, `www.contoso.com/addin2`), they are also permitted to share information back and forth through `storage`.</span></span> <span data-ttu-id="fc85f-120">Observe que os suplementos que possuem subdomínios diferentes terão instâncias diferentes `storage` (por exemplo, `subdomain.contoso.com/addin1` , `differentsubdomain.contoso.com/addin2` ).</span><span class="sxs-lookup"><span data-stu-id="fc85f-120">Note that add-ins which have different subdomains will have different instances of `storage` (for example, `subdomain.contoso.com/addin1`, `differentsubdomain.contoso.com/addin2`).</span></span>

<span data-ttu-id="fc85f-121">Como `storage` pode ser um local compartilhado, é importante notar que é possível substituir os pares chave-valor.</span><span class="sxs-lookup"><span data-stu-id="fc85f-121">Because `storage` can be a shared location, it is important to realize that it is possible to override key-value pairs.</span></span>

<span data-ttu-id="fc85f-122">Os métodos a seguir estão disponíveis no objeto `storage`:</span><span class="sxs-lookup"><span data-stu-id="fc85f-122">The following methods are available on the `storage` object:</span></span>

 - `getItem`
 - `getItems`
 - `setItem`
 - `setItems`
 - `removeItem`
 - `removeItems`
 - `getKeys`

<span data-ttu-id="fc85f-123">.</span><span class="sxs-lookup"><span data-stu-id="fc85f-123">.</span></span>[!NOTE]
> <span data-ttu-id="fc85f-124">Não há nenhum método para limpar todas as informações (como `clear` ).</span><span class="sxs-lookup"><span data-stu-id="fc85f-124">There's no method for clearing all information (such as `clear`).</span></span> <span data-ttu-id="fc85f-125">Em vez disso, use `removeItems` para remover várias entradas de uma só vez.</span><span class="sxs-lookup"><span data-stu-id="fc85f-125">Instead, you should instead use `removeItems` to remove multiple entries at a time.</span></span>

### <a name="officeruntimestorage-example"></a><span data-ttu-id="fc85f-126">Exemplo de OfficeRuntime. Storage</span><span class="sxs-lookup"><span data-stu-id="fc85f-126">OfficeRuntime.storage example</span></span>

<span data-ttu-id="fc85f-127">O exemplo de código a seguir chama a `OfficeRuntime.storage.setItem` função para definir uma chave e um valor para `storage` .</span><span class="sxs-lookup"><span data-stu-id="fc85f-127">The following code sample calls the `OfficeRuntime.storage.setItem` function to set a key and value into `storage`.</span></span>

```js
function StoreValue(key, value) {

  return OfficeRuntime.storage.setItem(key, value).then(function (result) {
      return "Success: Item with key '" + key + "' saved to storage.";
  }, function (error) {
      return "Error: Unable to save item with key '" + key + "' to storage. " + error;
  });
}
```

## <a name="additional-considerations"></a><span data-ttu-id="fc85f-128">Considerações adicionais</span><span class="sxs-lookup"><span data-stu-id="fc85f-128">Additional considerations</span></span>

<span data-ttu-id="fc85f-129">Se o suplemento usar apenas funções personalizadas sem interface do usuário, observe que não é possível acessar o modelo de objeto de documento (DOM) com funções personalizadas sem interface do usuário ou usar bibliotecas como jQuery que dependem do DOM.</span><span class="sxs-lookup"><span data-stu-id="fc85f-129">If your add-in only uses UI-less custom functions, note that you can't access the Document Object Model (DOM) with UI-less custom functions or use libraries like jQuery that rely on the DOM.</span></span>

## <a name="next-steps"></a><span data-ttu-id="fc85f-130">Próximas etapas</span><span class="sxs-lookup"><span data-stu-id="fc85f-130">Next steps</span></span>
<span data-ttu-id="fc85f-131">Saiba como [depurar funções personalizadas sem interface do usuário](custom-functions-debugging.md).</span><span class="sxs-lookup"><span data-stu-id="fc85f-131">Learn how to [debug UI-less custom functions](custom-functions-debugging.md).</span></span>

## <a name="see-also"></a><span data-ttu-id="fc85f-132">Confira também</span><span class="sxs-lookup"><span data-stu-id="fc85f-132">See also</span></span>

* [<span data-ttu-id="fc85f-133">Autenticar funções personalizadas sem interface do usuário</span><span class="sxs-lookup"><span data-stu-id="fc85f-133">Authenticate UI-less custom functions</span></span>](custom-functions-authentication.md)
* [<span data-ttu-id="fc85f-134">Criar funções personalizadas no Excel</span><span class="sxs-lookup"><span data-stu-id="fc85f-134">Create custom functions in Excel</span></span>](custom-functions-overview.md)
* [<span data-ttu-id="fc85f-135">Tutorial de funções personalizadas</span><span class="sxs-lookup"><span data-stu-id="fc85f-135">Custom functions tutorial</span></span>](../tutorials/excel-tutorial-create-custom-functions.md)
