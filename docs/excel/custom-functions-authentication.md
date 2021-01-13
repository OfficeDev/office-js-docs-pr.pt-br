---
ms.date: 05/17/2020
description: Autenticar usuários usando funções personalizadas no Excel que não usam o painel de tarefas.
title: Autenticação para funções personalizadas sem interface do usuário
localization_priority: Normal
ms.openlocfilehash: bca3cd422330b6499e18c31ef8d7da6def81b546
ms.sourcegitcommit: d28392721958555d6edea48cea000470bd27fcf7
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 01/13/2021
ms.locfileid: "49839856"
---
# <a name="authentication-for-ui-less-custom-functions"></a><span data-ttu-id="277e2-103">Autenticação para funções personalizadas sem interface do usuário</span><span class="sxs-lookup"><span data-stu-id="277e2-103">Authentication for UI-less custom functions</span></span>

<span data-ttu-id="277e2-104">Em alguns cenários, sua função personalizada que não usa um painel de tarefas ou outros elementos da interface do usuário (função personalizada sem interface do usuário) precisará autenticar o usuário para acessar recursos protegidos.</span><span class="sxs-lookup"><span data-stu-id="277e2-104">In some scenarios your custom function that does not use a task pane or other user interface elements (UI-less custom function) will need to authenticate the user in order to access protected resources.</span></span> <span data-ttu-id="277e2-105">Esteja ciente de que funções personalizadas sem interface do usuário são executados em um tempo de execução somente JavaScript.</span><span class="sxs-lookup"><span data-stu-id="277e2-105">Be aware that UI-less custom functions run in a JavaScript-only runtime.</span></span> <span data-ttu-id="277e2-106">Por isso, você precisará passar dados entre o tempo de execução somente JavaScript e o tempo de execução típico do mecanismo do navegador usado pela maioria dos complementos usando o objeto e a API de Caixa de `OfficeRuntime.storage` Diálogo.</span><span class="sxs-lookup"><span data-stu-id="277e2-106">Because of this, you'll need to pass data back and forth between the JavaScript-only runtime and the typical browser engine runtime used by most add-ins using the `OfficeRuntime.storage` object and the Dialog API.</span></span>

[!include[Excel custom functions note](../includes/excel-custom-functions-note.md)]

[!include[Shared runtime note](../includes/shared-runtime-note.md)]

## <a name="officeruntimestorage-object"></a><span data-ttu-id="277e2-107">Objeto OfficeRuntime.storage</span><span class="sxs-lookup"><span data-stu-id="277e2-107">OfficeRuntime.storage object</span></span>

<span data-ttu-id="277e2-108">O tempo de execução somente JavaScript usado por funções personalizadas sem interface do usuário não tem um objeto disponível na janela global, onde você normalmente `localStorage` armazena dados.</span><span class="sxs-lookup"><span data-stu-id="277e2-108">The JavaScript-only runtime used by UI-less custom functions doesn't have a `localStorage` object available on the global window, where you typically store data.</span></span> <span data-ttu-id="277e2-109">Em vez disso, você deve compartilhar dados entre funções personalizadas sem interface do usuário e painéis de tarefas usando [OfficeRuntime.storage](/javascript/api/office-runtime/officeruntime.storage) para definir e obter dados.</span><span class="sxs-lookup"><span data-stu-id="277e2-109">Instead, you should share data between UI-less custom functions and task panes by using [OfficeRuntime.storage](/javascript/api/office-runtime/officeruntime.storage) to set and get data.</span></span>

### <a name="suggested-usage"></a><span data-ttu-id="277e2-110">Uso sugerido</span><span class="sxs-lookup"><span data-stu-id="277e2-110">Suggested usage</span></span>

<span data-ttu-id="277e2-111">Quando você precisar se autenticar de uma função personalizada sem interface do usuário, verifique se o `storage` token de acesso já foi adquirido.</span><span class="sxs-lookup"><span data-stu-id="277e2-111">When you need to authenticate from a UI-less custom function, check `storage` to see if the access token was already acquired.</span></span> <span data-ttu-id="277e2-112">Caso contrário, use a API de caixa de diálogo para autenticar o usuário, recuperar o token de acesso e, em seguida, armazenar o token em `storage` para uso futuro.</span><span class="sxs-lookup"><span data-stu-id="277e2-112">If not, use the dialog API to authenticate the user, retrieve the access token, and then store the token in `storage` for future use.</span></span>

## <a name="dialog-api"></a><span data-ttu-id="277e2-113">API de Caixa de Diálogo</span><span class="sxs-lookup"><span data-stu-id="277e2-113">Dialog API</span></span>

<span data-ttu-id="277e2-114">Se um token não existir, você deverá usar a API de diálogo para solicitar que o usuário faça logon.</span><span class="sxs-lookup"><span data-stu-id="277e2-114">If a token doesn't exist, you should use the Dialog API to ask the user to sign in.</span></span> <span data-ttu-id="277e2-115">Depois que um usuário insere suas credenciais, o token de acesso resultante pode ser armazenado em `storage`.</span><span class="sxs-lookup"><span data-stu-id="277e2-115">After a user enters their credentials, the resulting access token can be stored in `storage`.</span></span>

> [!NOTE]
> <span data-ttu-id="277e2-116">O tempo de execução somente JavaScript usa um objeto Dialog ligeiramente diferente do objeto Dialog no tempo de execução do mecanismo do navegador usado pelos painéis de tarefas.</span><span class="sxs-lookup"><span data-stu-id="277e2-116">The JavaScript-only runtime uses a Dialog object that is slightly different from the Dialog object in the browser engine runtime used by task panes.</span></span> <span data-ttu-id="277e2-117">Ambos são chamados de "API da Caixa de Diálogo", mas são usadas para autenticar usuários no tempo de execução `OfficeRuntime.Dialog` somente JavaScript.</span><span class="sxs-lookup"><span data-stu-id="277e2-117">They're both referred to as the "Dialog API", but use `OfficeRuntime.Dialog` to authenticate users in the JavaScript-only runtime.</span></span>

<span data-ttu-id="277e2-118">O diagrama a seguir descreve esse processo básico.</span><span class="sxs-lookup"><span data-stu-id="277e2-118">The following diagram outlines this basic process.</span></span> <span data-ttu-id="277e2-119">A linha pontilhada indica que funções personalizadas sem interface do usuário e o painel de tarefas do seu complemento fazem parte do seu complemento como um todo, embora usem tempos de execução separados.</span><span class="sxs-lookup"><span data-stu-id="277e2-119">The dotted line indicates that UI-less custom functions and your add-in's task pane are both part of your add-in as a whole, though they use separate runtimes.</span></span>

1. <span data-ttu-id="277e2-120">Você emmitiu uma chamada de função personalizada sem interface do usuário de uma célula em uma planilha do Excel.</span><span class="sxs-lookup"><span data-stu-id="277e2-120">You issue a UI-less custom function call from a cell in an Excel workbook.</span></span>
2. <span data-ttu-id="277e2-121">A função personalizada sem interface do usuário usa `Dialog` para passar suas credenciais de usuário para um site.</span><span class="sxs-lookup"><span data-stu-id="277e2-121">The UI-less custom function uses `Dialog` to pass your user credentials to a website.</span></span>
3. <span data-ttu-id="277e2-122">Em seguida, este site retorna um token de acesso à função personalizada sem interface do usuário.</span><span class="sxs-lookup"><span data-stu-id="277e2-122">This website then returns an access token to the UI-less custom function.</span></span>
4. <span data-ttu-id="277e2-123">Sua função personalizada sem interface do usuário define esse token de acesso como `storage` .</span><span class="sxs-lookup"><span data-stu-id="277e2-123">Your UI-less custom function then sets this access token to the `storage`.</span></span>
5. <span data-ttu-id="277e2-124">O painel de tarefas do seu suplemento acessa o token a partir de `storage`.</span><span class="sxs-lookup"><span data-stu-id="277e2-124">Your add-in's task pane accesses the token from `storage`.</span></span>

<span data-ttu-id="277e2-125">![Diagrama de função personalizada usando a API de caixa de diálogo para obter o token de acesso e compartilhar o token com o painel de tarefas por meio da API OfficeRuntime.storage.](../images/authentication-diagram.png "Diagrama de autenticação.")</span><span class="sxs-lookup"><span data-stu-id="277e2-125">![Diagram of custom function using dialog API to get access token, and then share token with task pane through the OfficeRuntime.storage API.](../images/authentication-diagram.png "Authentication diagram.")</span></span>

## <a name="storing-the-token"></a><span data-ttu-id="277e2-126">Armazenando o token</span><span class="sxs-lookup"><span data-stu-id="277e2-126">Storing the token</span></span>

<span data-ttu-id="277e2-127">Os exemplos a seguir são do exemplo de código [Usando OfficeRuntime.storage em funções personalizadas](https://github.com/OfficeDev/PnP-OfficeAddins/tree/master/Excel-custom-functions/AsyncStorage).</span><span class="sxs-lookup"><span data-stu-id="277e2-127">The following examples are from the [Using OfficeRuntime.storage in custom functions](https://github.com/OfficeDev/PnP-OfficeAddins/tree/master/Excel-custom-functions/AsyncStorage) code sample.</span></span> <span data-ttu-id="277e2-128">Consulte este exemplo de código para ver um exemplo completo de compartilhamento de dados entre funções personalizadas sem interface do usuário e o painel de tarefas.</span><span class="sxs-lookup"><span data-stu-id="277e2-128">Refer to this code sample for a complete example of sharing data between UI-less custom functions and the task pane.</span></span>

<span data-ttu-id="277e2-129">Se a função personalizada sem interface do usuário autenticar, ela receberá o token de acesso e precisará armazená-lo. `storage`</span><span class="sxs-lookup"><span data-stu-id="277e2-129">If the UI-less custom function authenticates, then it receives the access token and will need to store it in `storage`.</span></span> <span data-ttu-id="277e2-130">O exemplo de código a seguir mostra como chamar o método `storage.setItem` para armazenar um valor.</span><span class="sxs-lookup"><span data-stu-id="277e2-130">The following code sample shows how to call the `storage.setItem` method to store a value.</span></span> <span data-ttu-id="277e2-131">A função é uma função personalizada sem interface do usuário que, por exemplo, armazena um `storeValue` valor do usuário.</span><span class="sxs-lookup"><span data-stu-id="277e2-131">The `storeValue` function is a UI-less custom function that for example purposes stores a value from the user.</span></span> <span data-ttu-id="277e2-132">Você pode modificá-la para que seja armazenado qualquer valor de token que você precise.</span><span class="sxs-lookup"><span data-stu-id="277e2-132">You can modify this to store any token value you need.</span></span>

```js
/**
 * Stores a key-value pair into OfficeRuntime.storage.
 * @customfunction
 * @param {string} key Key of item to put into storage.
 * @param {*} value Value of item to put into storage.
 */
function storeValue(key, value) {
  return OfficeRuntime.storage.setItem(key, value).then(function (result) {
      return "Success: Item with key '" + key + "' saved to storage.";
  }, function (error) {
      return "Error: Unable to save item with key '" + key + "' to storage. " + error;
  });
}
```

<span data-ttu-id="277e2-133">Quando o painel de tarefas precisa do token de acesso, ele pode recuperar o token de `storage`.</span><span class="sxs-lookup"><span data-stu-id="277e2-133">When the task pane needs the access token, it can retrieve the token from `storage`.</span></span> <span data-ttu-id="277e2-134">O exemplo de código a seguir mostra como usar o método `storage.getItem` para recuperar o token.</span><span class="sxs-lookup"><span data-stu-id="277e2-134">The following code sample shows how to use the `storage.getItem` method to retrieve the token.</span></span>

```js
/**
 * Read a token from storage.
 * @customfunction GETTOKEN
 */
function receiveTokenFromCustomFunction() {
  var key = "token";
  var tokenSendStatus = document.getElementById('tokenSendStatus');
  OfficeRuntime.storage.getItem(key).then(function (result) {
     tokenSendStatus.value = "Success: Item with key '" + key + "' read from storage.";
     document.getElementById('tokenTextBox2').value = result;
  }, function (error) {
     tokenSendStatus.value = "Error: Unable to read item with key '" + key + "' from storage. " + error;
  });
}
```

## <a name="general-guidance"></a><span data-ttu-id="277e2-135">Orientação geral</span><span class="sxs-lookup"><span data-stu-id="277e2-135">General guidance</span></span>

<span data-ttu-id="277e2-136">Os Suplementos do Office são baseados na Web e você pode usar qualquer técnica de autenticação da Web.</span><span class="sxs-lookup"><span data-stu-id="277e2-136">Office Add-ins are web-based and you can use any web authentication technique.</span></span> <span data-ttu-id="277e2-137">Não há nenhum padrão ou método específico que você deve seguir para implementar sua própria autenticação com funções personalizadas sem interface do usuário.</span><span class="sxs-lookup"><span data-stu-id="277e2-137">There is no particular pattern or method you must follow to implement your own authentication with UI-less custom functions.</span></span> <span data-ttu-id="277e2-138">Você pode querer consultar a documentação sobre vários padrões de autenticação, começando com [este artigo sobre a autorização por serviços externos](../develop/auth-external-add-ins.md).</span><span class="sxs-lookup"><span data-stu-id="277e2-138">You may wish to consult the documentation about various authentication patterns, starting with [this article about authorizing via external services](../develop/auth-external-add-ins.md).</span></span>  

<span data-ttu-id="277e2-139">Evite usar os seguintes locais para armazenar dados ao desenvolver funções personalizadas:</span><span class="sxs-lookup"><span data-stu-id="277e2-139">Avoid using the following locations to store data when developing custom functions:</span></span>  

- <span data-ttu-id="277e2-140">`localStorage`: funções personalizadas sem interface do usuário não têm acesso ao objeto global e, portanto, não têm acesso `window` aos dados armazenados `localStorage` em .</span><span class="sxs-lookup"><span data-stu-id="277e2-140">`localStorage`: UI-less custom functions do not have access to the global `window` object and therefore have no access to data stored in `localStorage`.</span></span>
- <span data-ttu-id="277e2-141">`Office.context.document.settings`: Esse local não é seguro, e informações podem ser extraídas por qualquer pessoa usando o suplemento.</span><span class="sxs-lookup"><span data-stu-id="277e2-141">`Office.context.document.settings`:  This location is not secure and information can be extracted by anyone using the add-in.</span></span>

## <a name="dialog-box-api-example"></a><span data-ttu-id="277e2-142">Exemplo de API da caixa de diálogo</span><span class="sxs-lookup"><span data-stu-id="277e2-142">Dialog box API example</span></span>

<span data-ttu-id="277e2-143">No exemplo de código a seguir, a função usa a `getTokenViaDialog` função da API para exibir uma caixa de `Dialog` `displayWebDialogOptions` diálogo.</span><span class="sxs-lookup"><span data-stu-id="277e2-143">In the following code sample, the function `getTokenViaDialog` uses the `Dialog` API's `displayWebDialogOptions` function to display a dialog box.</span></span> <span data-ttu-id="277e2-144">Este exemplo é fornecido para mostrar os recursos do `Dialog` objeto, não demonstrar como autenticar.</span><span class="sxs-lookup"><span data-stu-id="277e2-144">This sample is provided to show the capabilities of the `Dialog` object, not demonstrate how to authenticate.</span></span>

```JavaScript
/**
 * Function retrieves a cached token or opens a dialog box if there is no saved token. Note that this is not a sufficient example of authentication but is intended to show the capabilities of the Dialog object.
 * @param {string} url URL for a stored token.
 */
function getTokenViaDialog(url) {
  return new Promise (function (resolve, reject) {
    if (_dialogOpen) {
      // Can only have one dialog box open at once. Wait for previous dialog box's token.
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
          dialog.close();
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
```

## <a name="next-steps"></a><span data-ttu-id="277e2-145">Próximas etapas</span><span class="sxs-lookup"><span data-stu-id="277e2-145">Next steps</span></span>
<span data-ttu-id="277e2-146">Saiba como [depurar funções personalizadas sem](custom-functions-debugging.md)interface do usuário.</span><span class="sxs-lookup"><span data-stu-id="277e2-146">Learn how to [debug UI-less custom functions](custom-functions-debugging.md).</span></span>

## <a name="see-also"></a><span data-ttu-id="277e2-147">Confira também</span><span class="sxs-lookup"><span data-stu-id="277e2-147">See also</span></span>

* [<span data-ttu-id="277e2-148">Tempo de execução para funções personalizadas do Excel sem interface do usuário</span><span class="sxs-lookup"><span data-stu-id="277e2-148">Runtime for UI-less Excel custom functions</span></span>](custom-functions-runtime.md)
* [<span data-ttu-id="277e2-149">Tutorial de funções personalizadas do Excel</span><span class="sxs-lookup"><span data-stu-id="277e2-149">Excel custom functions tutorial</span></span>](../tutorials/excel-tutorial-create-custom-functions.md)