---
ms.date: 04/15/2019
description: Autenticar usuários usando funções personalizadas no Excel.
title: Autenticação para funções personalizadas
ms.openlocfilehash: 75ffb82c0dc9350c35b22b1d1676990598ea0c44
ms.sourcegitcommit: 9e7b4daa8d76c710b9d9dd4ae2e3c45e8fe07127
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 04/24/2019
ms.locfileid: "32449314"
---
# <a name="authentication"></a><span data-ttu-id="82311-103">Autenticação</span><span class="sxs-lookup"><span data-stu-id="82311-103">Authentication</span></span>

<span data-ttu-id="82311-104">Em alguns cenários, a função personalizada precisará autenticar o usuário para poder acessar recursos protegidos.</span><span class="sxs-lookup"><span data-stu-id="82311-104">In some scenarios your custom function will need to authenticate the user in order to access protected resources.</span></span> <span data-ttu-id="82311-105">Embora as funções personalizadas não exijam um método de autenticação específico, você deve estar ciente de que as funções personalizadas são executadas em um tempo de execução separado do painel de tarefas e de outros elementos de interface do usuário do seu suplemento.</span><span class="sxs-lookup"><span data-stu-id="82311-105">While custom functions don't require a specific method of authentication, you should be aware that custom functions run in a separate runtime from the task pane and other UI elements of your add-in.</span></span> <span data-ttu-id="82311-106">Por causa disso, você precisará transmitir dados entre os dois tempos de execução usando o `AsyncStorage` objeto e a API da caixa de diálogo.</span><span class="sxs-lookup"><span data-stu-id="82311-106">Because of this, you'll need to pass data back and forth between the two runtimes using the `AsyncStorage` object and the Dialog API.</span></span>
  
## <a name="asyncstorage-object"></a><span data-ttu-id="82311-107">Objeto AsyncStorage</span><span class="sxs-lookup"><span data-stu-id="82311-107">AsyncStorage object</span></span>

<span data-ttu-id="82311-108">O tempo de execução de funções personalizadas `localStorage` não tem um objeto disponível na janela global, onde você normalmente pode armazenar dados.</span><span class="sxs-lookup"><span data-stu-id="82311-108">The custom functions runtime doesn't have a `localStorage` object available on the global window, where you might typically store data.</span></span> <span data-ttu-id="82311-109">Em vez disso, você deve compartilhar dados entre funções personalizadas e painéis de tarefas usando o [OfficeRuntime. AsyncStorage](/javascript/api/office-runtime/officeruntime.asyncstorage) para definir e obter dados.</span><span class="sxs-lookup"><span data-stu-id="82311-109">Instead, you should share data between custom functions and task panes by using [OfficeRuntime.AsyncStorage](/javascript/api/office-runtime/officeruntime.asyncstorage) to set and get data.</span></span>

<span data-ttu-id="82311-110">Além disso, há um benefício em usar `AsyncStorage`o; Ele usa um ambiente de área restrita seguro para que seus dados não possam ser acessados por outros suplementos.</span><span class="sxs-lookup"><span data-stu-id="82311-110">Additionally, there is a benefit to using `AsyncStorage`; it uses a secure sandbox environment so that your data cannot be accessed by other add-ins.</span></span>

### <a name="suggested-usage"></a><span data-ttu-id="82311-111">Uso sugerido</span><span class="sxs-lookup"><span data-stu-id="82311-111">Suggested usage</span></span>

<span data-ttu-id="82311-112">Quando você precisar autenticar do painel de tarefas ou de uma função personalizada, verifique `AsyncStorage` se o token de acesso já foi adquirido.</span><span class="sxs-lookup"><span data-stu-id="82311-112">When you need to authenticate either from the task pane or a custom function, check `AsyncStorage` to see if the access token was already acquired.</span></span> <span data-ttu-id="82311-113">Caso contrário, use a API de caixa de diálogo para autenticar o usuário, recuperar o token de acesso e armazená-lo `AsyncStorage` para uso futuro.</span><span class="sxs-lookup"><span data-stu-id="82311-113">If not, use the dialog API to authenticate the user, retrieve the access token, and then store the token in `AsyncStorage` for future use.</span></span>

## <a name="dialog-api"></a><span data-ttu-id="82311-114">API da caixa de diálogo</span><span class="sxs-lookup"><span data-stu-id="82311-114">Dialog API</span></span>

<span data-ttu-id="82311-115">Se não houver um token, você deverá usar a API da caixa de diálogo para solicitar que o usuário entre.</span><span class="sxs-lookup"><span data-stu-id="82311-115">If a token doesn't exist, you should use the Dialog API to ask the user to sign in.</span></span> <span data-ttu-id="82311-116">Após um usuário inserir suas credenciais, o token de acesso resultante poderá ser armazenado `AsyncStorage`no.</span><span class="sxs-lookup"><span data-stu-id="82311-116">After a user enters their credentials, the resulting access token can be stored in `AsyncStorage`.</span></span>

> [!NOTE]
> <span data-ttu-id="82311-117">O tempo de execução de funções personalizadas usa um objeto Dialog que é ligeiramente diferente do objeto Dialog no tempo de execução do mecanismo de navegador usado por painéis de tarefas.</span><span class="sxs-lookup"><span data-stu-id="82311-117">The custom functions runtime uses a Dialog object that is slightly different from the Dialog object in the browser engine runtime used by task panes.</span></span> <span data-ttu-id="82311-118">Eles são conhecidos como "API da caixa de diálogo", mas usam `Officeruntime.Dialog` para autenticar usuários no tempo de execução de funções personalizadas.</span><span class="sxs-lookup"><span data-stu-id="82311-118">They're both referred to as the "Dialog API", but use `Officeruntime.Dialog` to authenticate users in the custom functions runtime.</span></span>

<span data-ttu-id="82311-119">Para obter informações sobre como usar a `OfficeRuntime.Dialog`caixa de [diálogo, consulte funções personalizadas](/office/dev/add-ins/excel/custom-functions-dialog).</span><span class="sxs-lookup"><span data-stu-id="82311-119">For information on how to use the `OfficeRuntime.Dialog`, see [Custom Functions dialog](/office/dev/add-ins/excel/custom-functions-dialog).</span></span>

<span data-ttu-id="82311-120">Ao planejar todo o processo de autenticação como um todo, talvez seja útil pensar no painel de tarefas e nos elementos de interface do usuário do suplemento e das funções personalizadas, que fazem parte do suplemento como entidades separadas que podem se comunicar entre si `AsyncStorage`.</span><span class="sxs-lookup"><span data-stu-id="82311-120">When envisioning the entire authentication process as a whole, it might be helpful to think of the task pane and UI elements of your add-in and the custom functions part of your add-in as separate entities which can communicate with each other through `AsyncStorage`.</span></span>

<span data-ttu-id="82311-121">O diagrama a seguir descreve esse processo básico.</span><span class="sxs-lookup"><span data-stu-id="82311-121">The following diagram outlines this basic process.</span></span> <span data-ttu-id="82311-122">Observe que a linha pontilhada indica que, enquanto elas executam ações separadas, as funções personalizadas e o painel de tarefas do seu suplemento fazem parte do seu suplemento como um todo.</span><span class="sxs-lookup"><span data-stu-id="82311-122">Note that the dotted line indicates that while they perform separate actions, custom functions and your add-in's task pane are both part of your add-in as a whole.</span></span>

1. <span data-ttu-id="82311-123">Você emite uma chamada de função personalizada a partir de uma célula em uma pasta de trabalho do Excel.</span><span class="sxs-lookup"><span data-stu-id="82311-123">You issue a custom function call from a cell in an Excel workbook.</span></span>
2. <span data-ttu-id="82311-124">A função personalizada usa `Officeruntime.Dialog` o para passar suas credenciais de usuário para um site.</span><span class="sxs-lookup"><span data-stu-id="82311-124">The custom function uses `Officeruntime.Dialog` to pass your user credentials to a website.</span></span>
3. <span data-ttu-id="82311-125">Este site, em seguida, retorna um token de acesso para a função personalizada.</span><span class="sxs-lookup"><span data-stu-id="82311-125">This website then returns an access token to the custom function.</span></span>
4. <span data-ttu-id="82311-126">Sua função personalizada então define esse token de acesso para `AsyncStorage`o.</span><span class="sxs-lookup"><span data-stu-id="82311-126">Your custom function then sets this access token to the `AsyncStorage`.</span></span>
5. <span data-ttu-id="82311-127">O painel de tarefas do suplemento acessa o token de `AsyncStorage`.</span><span class="sxs-lookup"><span data-stu-id="82311-127">Your add-in's task pane accesses the token from `AsyncStorage`.</span></span>

<span data-ttu-id="82311-128">![Diagrama da função personalizada usando a API da caixa de diálogo para obter o token de acesso e compartilhar o token com o painel de tarefas por meio da API AsyncStorage.] (../images/authentication-diagram.png "Diagrama de autenticação.")</span><span class="sxs-lookup"><span data-stu-id="82311-128">![Diagram of custom function using dialog API to get access token, and then share token with task pane through the AsyncStorage API.](../images/authentication-diagram.png "Authentication diagram.")</span></span>

## <a name="storing-the-token"></a><span data-ttu-id="82311-129">Armazenar o token</span><span class="sxs-lookup"><span data-stu-id="82311-129">Storing the token</span></span>

<span data-ttu-id="82311-130">Os exemplos a seguir são do [usando AsyncStorage no exemplo de código de funções personalizadas](https://github.com/OfficeDev/PnP-OfficeAddins/tree/master/Excel-custom-functions/AsyncStorage) .</span><span class="sxs-lookup"><span data-stu-id="82311-130">The following examples are from the [Using AsyncStorage in custom functions](https://github.com/OfficeDev/PnP-OfficeAddins/tree/master/Excel-custom-functions/AsyncStorage) code sample.</span></span> <span data-ttu-id="82311-131">Consulte este exemplo de código para obter um exemplo completo de compartilhamento de dados entre funções personalizadas e o painel de tarefas.</span><span class="sxs-lookup"><span data-stu-id="82311-131">Refer to this code sample for a complete example of sharing data between custom functions and the task pane.</span></span>

<span data-ttu-id="82311-132">Se a função personalizada autenticar, ela receberá o token de acesso e deverá armazená-lo `AsyncStorage`no.</span><span class="sxs-lookup"><span data-stu-id="82311-132">If the custom function authenticates, then it receives the access token and will need to store it in `AsyncStorage`.</span></span> <span data-ttu-id="82311-133">O exemplo de código a seguir mostra como chamar `AsyncStorage.setItem` o método para armazenar um valor.</span><span class="sxs-lookup"><span data-stu-id="82311-133">The following code sample shows how to call the `AsyncStorage.setItem` method to store a value.</span></span> <span data-ttu-id="82311-134">A `StoreValue` função é uma função personalizada que, por exemplo, armazena um valor do usuário.</span><span class="sxs-lookup"><span data-stu-id="82311-134">The `StoreValue` function is a custom function that for example purposes stores a value from the user.</span></span> <span data-ttu-id="82311-135">Você pode modificá-lo para armazenar qualquer valor de token necessário.</span><span class="sxs-lookup"><span data-stu-id="82311-135">You can modify this to store any token value you need.</span></span>

```javascript
function StoreValue(key, value) {
  return OfficeRuntime.AsyncStorage.setItem(key, value).then(function (result) {
      return "Success: Item with key '" + key + "' saved to AsyncStorage.";
  }, function (error) {
      return "Error: Unable to save item with key '" + key + "' to AsyncStorage. " + error;
  });
}
```

<span data-ttu-id="82311-136">Quando o painel de tarefas precisa do token de acesso, ele pode recuperar o `AsyncStorage`token de.</span><span class="sxs-lookup"><span data-stu-id="82311-136">When the task pane needs the access token, it can retrieve the token from `AsyncStorage`.</span></span> <span data-ttu-id="82311-137">O exemplo de código a seguir mostra como usar `AsyncStorage.getItem` o método para recuperar o token.</span><span class="sxs-lookup"><span data-stu-id="82311-137">The following code sample shows how to use the `AsyncStorage.getItem` method to retrieve the token.</span></span>

```javascript
function ReceiveTokenFromCustomFunction() {
   var key = "token";
   var tokenSendStatus = document.getElementById('tokenSendStatus');
   OfficeRuntime.AsyncStorage.getItem(key).then(function (result) {
      tokenSendStatus.value = "Success: Item with key '" + key + "' read from AsyncStorage.";
      document.getElementById('tokenTextBox2').value = result;
   }, function (error) {
      tokenSendStatus.value = "Error: Unable to read item with key '" + key + "' from AsyncStorage. " + error;
   });
}
```

## <a name="general-guidance"></a><span data-ttu-id="82311-138">Orientação geral</span><span class="sxs-lookup"><span data-stu-id="82311-138">General guidance</span></span>

<span data-ttu-id="82311-139">Os suplementos do Office são baseados na Web e você pode usar qualquer técnica de autenticação da Web.</span><span class="sxs-lookup"><span data-stu-id="82311-139">Office Add-ins are web-based and you can use any web authentication technique.</span></span> <span data-ttu-id="82311-140">Não há um padrão ou método específico que você deve seguir para implementar sua própria autenticação com funções personalizadas.</span><span class="sxs-lookup"><span data-stu-id="82311-140">There is no particular pattern or method you must follow to implement your own authentication with custom functions.</span></span> <span data-ttu-id="82311-141">Você pode querer consultar a documentação sobre vários padrões de autenticação, começando com [Este artigo sobre como autorizar por meio de serviços externos](/office/dev/add-ins/develop/auth-external-add-ins?view=office-js).</span><span class="sxs-lookup"><span data-stu-id="82311-141">You may wish to consult the documentation about various authentication patterns, starting with [this article about authorizing via external services](/office/dev/add-ins/develop/auth-external-add-ins?view=office-js).</span></span>  

<span data-ttu-id="82311-142">Evite usar os seguintes locais para armazenar dados ao desenvolver funções personalizadas:</span><span class="sxs-lookup"><span data-stu-id="82311-142">Avoid using the following locations to store data when developing custom functions:</span></span>  

- <span data-ttu-id="82311-143">`localStorage`: As funções personalizadas não têm acesso ao objeto global `window` e, portanto, não têm acesso aos dados armazenados `localStorage`no.</span><span class="sxs-lookup"><span data-stu-id="82311-143">`localStorage`: Custom functions do not have access to the global `window` object and therefore have no access to data     stored in `localStorage`.</span></span>
- <span data-ttu-id="82311-144">`Office.context.document.settings`: Esse local não é seguro e as informações podem ser extraídas por qualquer pessoa que use o suplemento.</span><span class="sxs-lookup"><span data-stu-id="82311-144">`Office.context.document.settings`:  This location is not secure and information can be extracted by anyone using the     add-in.</span></span>

## <a name="see-also"></a><span data-ttu-id="82311-145">Confira também</span><span class="sxs-lookup"><span data-stu-id="82311-145">See also</span></span>

* [<span data-ttu-id="82311-146">Metadados de funções personalizadas</span><span class="sxs-lookup"><span data-stu-id="82311-146">Custom functions metadata</span></span>](custom-functions-json.md)
* [<span data-ttu-id="82311-147">Tempo de execução de funções personalizadas do Excel</span><span class="sxs-lookup"><span data-stu-id="82311-147">Runtime for Excel custom functions</span></span>](custom-functions-runtime.md)
* <span data-ttu-id="82311-148">[Práticas recomendadas de funções personalizadas](custom-functions-best-practices.md).</span><span class="sxs-lookup"><span data-stu-id="82311-148">[Custom functions best practices](custom-functions-best-practices.md)</span></span>
* [<span data-ttu-id="82311-149">Tutorial de funções personalizadas do Excel</span><span class="sxs-lookup"><span data-stu-id="82311-149">Excel custom functions tutorial</span></span>](excel-tutorial-custom-functions.md)
