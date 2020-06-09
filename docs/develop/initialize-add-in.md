---
title: Inicialize seu suplemento do Office
description: Saiba como inicializar o suplemento do Office.
ms.date: 02/27/2020
localization_priority: Normal
ms.openlocfilehash: 8310c5efb803391f7f0d4b01fda70dc0df537b21
ms.sourcegitcommit: be23b68eb661015508797333915b44381dd29bdb
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 06/08/2020
ms.locfileid: "44608136"
---
# <a name="initialize-your-office-add-in"></a><span data-ttu-id="7541d-103">Inicialize seu suplemento do Office</span><span class="sxs-lookup"><span data-stu-id="7541d-103">Initialize your Office Add-in</span></span>

<span data-ttu-id="7541d-104">Os Suplementos do Office têm sempre uma lógica de inicialização para fazer coisas como:</span><span class="sxs-lookup"><span data-stu-id="7541d-104">Office Add-ins often have start-up logic to do things such as:</span></span>

- <span data-ttu-id="7541d-105">Verifique se a versão do Office do usuário é compatível com todas as APIs do Office que seu código chama.</span><span class="sxs-lookup"><span data-stu-id="7541d-105">Check that the user's version of Office supports all the Office APIs that your code calls.</span></span>

- <span data-ttu-id="7541d-106">Certifique-se de que a existência de determinados artefatos, como uma planilha com um nome específico.</span><span class="sxs-lookup"><span data-stu-id="7541d-106">Ensure the existence of certain artifacts, such as a worksheet with a specific name.</span></span>

- <span data-ttu-id="7541d-107">Solicita que o usuário selecione algumas células no Excel e, em seguida, insira um gráfico inicializado com os valores selecionados.</span><span class="sxs-lookup"><span data-stu-id="7541d-107">Prompt the user to select some cells in Excel, and then insert a chart initialized with those selected values.</span></span>

- <span data-ttu-id="7541d-108">Estabeleça associações.</span><span class="sxs-lookup"><span data-stu-id="7541d-108">Establish bindings.</span></span>

- <span data-ttu-id="7541d-109">Use a API de caixa de diálogo do Office para solicitar ao usuário os valores padrão das configurações do suplemento.</span><span class="sxs-lookup"><span data-stu-id="7541d-109">Use the Office Dialog API to prompt the user for default add-in settings values.</span></span>

<span data-ttu-id="7541d-110">No entanto, um suplemento do Office não pode chamar com êxito nenhuma API JavaScript do Office até que a biblioteca seja carregada.</span><span class="sxs-lookup"><span data-stu-id="7541d-110">However, an Office Add-in cannot successfully call any Office JavaScript APIs until the library has been loaded.</span></span> <span data-ttu-id="7541d-111">Este artigo descreve as duas maneiras pelas quais o código pode garantir que a biblioteca tenha sido carregada:</span><span class="sxs-lookup"><span data-stu-id="7541d-111">This article describes the two ways your code can ensure that the library has been loaded:</span></span>

- <span data-ttu-id="7541d-112">Inicializar `Office.onReady()` .</span><span class="sxs-lookup"><span data-stu-id="7541d-112">Initialize with `Office.onReady()`.</span></span>
- <span data-ttu-id="7541d-113">Inicializar `Office.initialize` .</span><span class="sxs-lookup"><span data-stu-id="7541d-113">Initialize with `Office.initialize`.</span></span>

> [!TIP]
> <span data-ttu-id="7541d-114">Recomendamos que use `Office.onReady()`em vez de`Office.initialize`.</span><span class="sxs-lookup"><span data-stu-id="7541d-114">We recommend that you use `Office.onReady()` instead of `Office.initialize`.</span></span> <span data-ttu-id="7541d-115">Embora `Office.initialize` ainda tenha suporte, o `Office.onReady()` oferece mais flexibilidade.</span><span class="sxs-lookup"><span data-stu-id="7541d-115">Although `Office.initialize` is still supported, `Office.onReady()` provides more flexibility.</span></span> <span data-ttu-id="7541d-116">Você pode atribuir apenas um manipulador ao `Office.initialize` e ele é chamado apenas uma vez pela infraestrutura do Office.</span><span class="sxs-lookup"><span data-stu-id="7541d-116">You can assign only one handler to `Office.initialize` and it's called only once by the Office infrastructure.</span></span> <span data-ttu-id="7541d-117">Você pode chamar `Office.onReady()` diferentes locais no seu código e usar retornos de chamada diferentes.</span><span class="sxs-lookup"><span data-stu-id="7541d-117">You can call `Office.onReady()` in different places in your code and use different callbacks.</span></span>
> 
> <span data-ttu-id="7541d-118">Para saber mais sobre as diferenças entre essas técnicas, veja [Principais diferenças entre Office.initialize e Office.onReady()](#major-differences-between-officeinitialize-and-officeonready).</span><span class="sxs-lookup"><span data-stu-id="7541d-118">For information about the differences in these techniques, see [Major differences between Office.initialize and Office.onReady()](#major-differences-between-officeinitialize-and-officeonready).</span></span>

<span data-ttu-id="7541d-119">Para saber mais sobre a sequência de eventos na inicialização do suplemento, confira [Carregar o ambiente de tempo de execução e o DOM](loading-the-dom-and-runtime-environment.md).</span><span class="sxs-lookup"><span data-stu-id="7541d-119">For more details about the sequence of events when an add-in is initialized, see [Loading the DOM and runtime environment](loading-the-dom-and-runtime-environment.md).</span></span>

## <a name="initialize-with-officeonready"></a><span data-ttu-id="7541d-120">Inicializar com o Office.onReady()</span><span class="sxs-lookup"><span data-stu-id="7541d-120">Initialize with Office.onReady()</span></span>

<span data-ttu-id="7541d-121">`Office.onReady()`é um método assíncrono que retorna um objeto [Promise](https://developer.mozilla.org/docs/Web/JavaScript/Reference/Global_Objects/Promise) enquanto verifica se a biblioteca do Office. js foi carregada.</span><span class="sxs-lookup"><span data-stu-id="7541d-121">`Office.onReady()` is an asynchronous method that returns a [Promise](https://developer.mozilla.org/docs/Web/JavaScript/Reference/Global_Objects/Promise) object while it checks to see if the Office.js library is loaded.</span></span> <span data-ttu-id="7541d-122">Somente quando a biblioteca é carregada, ela resolve o Promise como um objeto que especifica o aplicativo host do Office com um valor de enumeração `Office.HostType` (`Excel`, `Word` etc.), e a plataforma com um valor de enumeração `Office.PlatformType` (`PC`, `Mac`, `OfficeOnline` etc.).</span><span class="sxs-lookup"><span data-stu-id="7541d-122">When the library is loaded, it resolves the Promise as an object that specifies the Office host application with an `Office.HostType` enum value (`Excel`, `Word`, etc.) and the platform with an `Office.PlatformType` enum value (`PC`, `Mac`, `OfficeOnline`, etc.).</span></span> <span data-ttu-id="7541d-123">O Promise será resolvido imediatamente quando a biblioteca estiver carregada ao`Office.onReady()` ser chamada.</span><span class="sxs-lookup"><span data-stu-id="7541d-123">The Promise resolves immediately if the library is already loaded when `Office.onReady()` is called.</span></span>

<span data-ttu-id="7541d-124">Uma maneira de chamar `Office.onReady()` é passá-la por um método de retorno de chamada.</span><span class="sxs-lookup"><span data-stu-id="7541d-124">One way to call `Office.onReady()` is to pass it a callback method.</span></span> <span data-ttu-id="7541d-125">Exemplo:</span><span class="sxs-lookup"><span data-stu-id="7541d-125">Here's an example:</span></span>

```js
Office.onReady(function(info) {
    if (info.host === Office.HostType.Excel) {
        // Do Excel-specific initialization (for example, make add-in task pane's
        // appearance compatible with Excel "green").
    }
    if (info.platform === Office.PlatformType.PC) {
        // Make minor layout changes in the task pane.
    }
    console.log(`Office.js is now ready in ${info.host} on ${info.platform}`);
});
```

<span data-ttu-id="7541d-126">Como alternativa, é possível encadear um método `then()` à chamada de `Office.onReady()`, em vez de passar um retorno de chamada.</span><span class="sxs-lookup"><span data-stu-id="7541d-126">Alternatively, you can chain a `then()` method to the call of `Office.onReady()`, instead of passing a callback.</span></span> <span data-ttu-id="7541d-127">Por exemplo, o código a seguir verifica se a versão do Excel do usuário é compatível com todas as APIs que o suplemento pode chamar.</span><span class="sxs-lookup"><span data-stu-id="7541d-127">For example, the following code checks to see that the user's version of Excel supports all the APIs that the add-in might call.</span></span>

```js
Office.onReady()
    .then(function() {
        if (!Office.context.requirements.isSetSupported('ExcelApi', '1.7')) {
            console.log("Sorry, this add-in only works with newer versions of Excel.");
        }
    });
```

<span data-ttu-id="7541d-128">Este é o mesmo exemplo que usa as palavras-chave `async` e `await` em TypeScript:</span><span class="sxs-lookup"><span data-stu-id="7541d-128">Here is the same example using the `async` and `await` keywords in TypeScript:</span></span>

```typescript
(async () => {
    await Office.onReady();
    if (!Office.context.requirements.isSetSupported('ExcelApi', '1.7')) {
        console.log("Sorry, this add-in only works with newer versions of Excel.");
    }
})();
```

<span data-ttu-id="7541d-129">Se estiver usando estruturas JavaScript adicionais que incluam testes e manipuladores próprios de inicialização, *geralmente* eles devem ser colocados dentro da resposta para `Office.onReady()`.</span><span class="sxs-lookup"><span data-stu-id="7541d-129">If you are using additional JavaScript frameworks that include their own initialization handler or tests, these should *usually* be placed within the response to `Office.onReady()`.</span></span> <span data-ttu-id="7541d-130">Por exemplo, a função [JQuery](https://jquery.com) `$(document).ready()` seria referenciada da seguinte maneira:</span><span class="sxs-lookup"><span data-stu-id="7541d-130">For example, [JQuery's](https://jquery.com) `$(document).ready()` function would be referenced as follows:</span></span>

```js
Office.onReady(function() {
    // Office is ready
    $(document).ready(function () {
        // The document is ready
    });
});
```

<span data-ttu-id="7541d-131">No entanto, há exceções a essa prática.</span><span class="sxs-lookup"><span data-stu-id="7541d-131">However, there are exceptions to this practice.</span></span> <span data-ttu-id="7541d-132">Por exemplo, digamos que você queira abrir o suplemento em um navegador (em vez de fazer sideload em um host do Office) para depurar a interface do usuário com ferramentas de navegador.</span><span class="sxs-lookup"><span data-stu-id="7541d-132">For example, suppose you want to open your add-in in a browser (instead of sideload it in an Office host) in order to debug your UI with browser tools.</span></span> <span data-ttu-id="7541d-133">Já que o Office.js não será carregado no navegador, `onReady` não será executado e o `$(document).ready` não será executado quando chamado dentro de `onReady` no Office.</span><span class="sxs-lookup"><span data-stu-id="7541d-133">Since Office.js won't load in the browser, `onReady` won't run and the `$(document).ready` won't run if it's called inside the Office `onReady`.</span></span> 

<span data-ttu-id="7541d-134">Outra exceção seria se você quiser que um indicador de progresso seja exibido no painel de tarefas enquanto o suplemento estiver sendo carregado.</span><span class="sxs-lookup"><span data-stu-id="7541d-134">Another exception would be if you want a progress indicator to appear in the task pane while the add-in is loading.</span></span> <span data-ttu-id="7541d-135">Neste cenário, o código deve chamar o jQuery `ready` e usar seu retorno de chamada para renderizar o indicador de progresso.</span><span class="sxs-lookup"><span data-stu-id="7541d-135">In this scenario, your code should call the jQuery `ready` and use its callback to render the progress indicator.</span></span> <span data-ttu-id="7541d-136">Em seguida, a chamada de retorno do Office `onReady` pode substituir o indicador de progresso com a interface do usuário final.</span><span class="sxs-lookup"><span data-stu-id="7541d-136">Then the Office `onReady`'s callback can replace the progress indicator with the final UI.</span></span> 

## <a name="initialize-with-officeinitialize"></a><span data-ttu-id="7541d-137">Inicializar com Office.initialize</span><span class="sxs-lookup"><span data-stu-id="7541d-137">Initialize with Office.initialize</span></span>

<span data-ttu-id="7541d-138">Um evento de inicialização é disparado quando a biblioteca do Office.js está carregada e pronta para a interação com o usuário.</span><span class="sxs-lookup"><span data-stu-id="7541d-138">An initialize event fires when the Office.js library is loaded and ready for user interaction.</span></span> <span data-ttu-id="7541d-139">É possível atribuir um manipulador ao `Office.initialize` que implementa a lógica de inicialização.</span><span class="sxs-lookup"><span data-stu-id="7541d-139">You can assign a handler to `Office.initialize` that implements your initialization logic.</span></span> <span data-ttu-id="7541d-140">Veja a seguir um exemplo que verifica se a versão do Excel do usuário é compatível com todas as APIs que o suplemento pode chamar.</span><span class="sxs-lookup"><span data-stu-id="7541d-140">The following is an example that checks to see that the user's version of Excel supports all the APIs that the add-in might call.</span></span>

```js
Office.initialize = function () {
    if (!Office.context.requirements.isSetSupported('ExcelApi', '1.7')) {
        console.log("Sorry, this add-in only works with newer versions of Excel.");
    }
};
```

<span data-ttu-id="7541d-141">Se você estiver usando estruturas JavaScript adicionais que incluam seu próprio manipulador de inicialização ou testes, elas *deverão ser* colocadas no `Office.initialize` evento (as exceções descritas na seção **inicializar com Office. onReady ()** anteriormente também serão aplicadas neste caso).</span><span class="sxs-lookup"><span data-stu-id="7541d-141">If you are using additional JavaScript frameworks that include their own initialization handler or tests, these should *usually* be placed within the `Office.initialize` event (the exceptions described in the **Initialize with Office.onReady()** section earlier apply in this case also).</span></span> <span data-ttu-id="7541d-142">Por exemplo, a função [JQuery](https://jquery.com) `$(document).ready()` seria referenciada da seguinte maneira:</span><span class="sxs-lookup"><span data-stu-id="7541d-142">For example, [JQuery's](https://jquery.com) `$(document).ready()` function would be referenced as follows:</span></span>

```js
Office.initialize = function () {
    // Office is ready
    $(document).ready(function () {
        // The document is ready
    });
  };
```

<span data-ttu-id="7541d-143">Para suplementos de conteúdo e painel de tarefas, `Office.initialize` fornece um parâmetro _reason_ adicional.</span><span class="sxs-lookup"><span data-stu-id="7541d-143">For task pane and content add-ins, `Office.initialize` provides an additional _reason_ parameter.</span></span> <span data-ttu-id="7541d-144">Esse parâmetro especifica como um suplemento foi adicionado ao documento atual.</span><span class="sxs-lookup"><span data-stu-id="7541d-144">This parameter specifies how an add-in was added to the current document.</span></span> <span data-ttu-id="7541d-145">Você pode usar isso para fornecer uma lógica diferente para quando um suplemento é inserido pela primeira vez, em comparação com quando já existia dentro do documento.</span><span class="sxs-lookup"><span data-stu-id="7541d-145">You can use this to provide different logic for when an add-in is first inserted versus when it already existed within the document.</span></span>

```js
Office.initialize = function (reason) {
    $(document).ready(function () {
        switch (reason) {
            case 'inserted': console.log('The add-in was just inserted.');
            case 'documentOpened': console.log('The add-in is already part of the document.');
        }
    });
 };
```

<span data-ttu-id="7541d-146">Para saber mais, veja [Evento Office.initialize](/javascript/api/office) e [Enumeração da InitializationReason](/javascript/api/office/office.initializationreason).</span><span class="sxs-lookup"><span data-stu-id="7541d-146">For more information, see [Office.initialize Event](/javascript/api/office) and [InitializationReason Enumeration](/javascript/api/office/office.initializationreason).</span></span>

## <a name="major-differences-between-officeinitialize-and-officeonready"></a><span data-ttu-id="7541d-147">Principais diferenças entre Office.initialize e Office.onReady</span><span class="sxs-lookup"><span data-stu-id="7541d-147">Major differences between Office.initialize and Office.onReady</span></span>

- <span data-ttu-id="7541d-148">É possível atribuir apenas um manipulador a `Office.initialize`, e ela é chamada apenas uma vez pela infraestrutura do Office, mas você pode chamar `Office.onReady()` em diferentes locais no código, e usar diferentes retornos de chamadas.</span><span class="sxs-lookup"><span data-stu-id="7541d-148">You can assign only one handler to `Office.initialize` and it's called only once by the Office infrastructure; but you can call `Office.onReady()` in different places in your code and use different callbacks.</span></span> <span data-ttu-id="7541d-149">Por exemplo, o código pode chamar `Office.onReady()`, logo que o script personalizado é carregado com um retorno de chamada que executa uma lógica de inicialização. Além disso, o código pode ter um botão no painel de tarefas, cujo script chama `Office.onReady()` com um retorno de chamada diferente.</span><span class="sxs-lookup"><span data-stu-id="7541d-149">For example, your code could call `Office.onReady()` as soon as your custom script loads with a callback that runs initialization logic; and your code could also have a button in the task pane, whose script calls `Office.onReady()` with a different callback.</span></span> <span data-ttu-id="7541d-150">Quando isso ocorre, o segundo retorno de chamada é executado quando o botão é clicado.</span><span class="sxs-lookup"><span data-stu-id="7541d-150">If so, the second callback runs when the button is clicked.</span></span>

- <span data-ttu-id="7541d-151">O evento `Office.initialize` é disparado no final do processo interno, e que o Office.js é inicializado automaticamente.</span><span class="sxs-lookup"><span data-stu-id="7541d-151">The `Office.initialize` event fires at the end of the internal process in which Office.js initializes itself.</span></span> <span data-ttu-id="7541d-152">Ele também é disparado *imediatamente* após o término do processo interno.</span><span class="sxs-lookup"><span data-stu-id="7541d-152">And it fires *immediately* after the internal process ends.</span></span> <span data-ttu-id="7541d-153">Se o código no qual você atribui um manipulador ao evento for executado muito tempo após o evento ser disparado, então o manipulador não será executado.</span><span class="sxs-lookup"><span data-stu-id="7541d-153">If the code in which you assign a handler to the event executes too long after the event fires, then your handler doesn't run.</span></span> <span data-ttu-id="7541d-154">Por exemplo, se estiver usando o gerenciador de tarefas WebPack, ele poderá configurar a home page do suplemento para carregar arquivos de polyfill, após carregar o Office.js, mas antes de carregar o JavaScript personalizado.</span><span class="sxs-lookup"><span data-stu-id="7541d-154">For example, if you are using the WebPack task manager, it might configure the add-in's home page to load polyfill files after it loads Office.js but before it loads your custom JavaScript.</span></span> <span data-ttu-id="7541d-155">Quando o script carrega e atribui o manipulador, o evento de inicialização já ocorreu.</span><span class="sxs-lookup"><span data-stu-id="7541d-155">By the time your script loads and assigns the handler, the initialize event has already happened.</span></span> <span data-ttu-id="7541d-156">Mas nunca é "tarde demais" para chamar `Office.onReady()`.</span><span class="sxs-lookup"><span data-stu-id="7541d-156">But it is never "too late" to call `Office.onReady()`.</span></span> <span data-ttu-id="7541d-157">Caso o evento de inicialização já tenha ocorrido, o retorno de chamada é executado imediatamente.</span><span class="sxs-lookup"><span data-stu-id="7541d-157">If the initialize event has already happened, the callback runs immediately.</span></span>

> [!NOTE]
> <span data-ttu-id="7541d-158">Mesmo que não tenha uma lógica de inicialização, você deve atribuir ou chamar `Office.onReady()` uma função vazia para `Office.initialize` quando o JavaScript do suplemento for carregado.</span><span class="sxs-lookup"><span data-stu-id="7541d-158">Even if you have no start-up logic, you should either call `Office.onReady()` or assign an empty function to `Office.initialize` when your add-in JavaScript loads.</span></span> <span data-ttu-id="7541d-159">Algumas combinações de host e da plataforma do Office não carregam o painel de tarefas até uma das delas aconteça.</span><span class="sxs-lookup"><span data-stu-id="7541d-159">Some Office host and platform combinations won't load the task pane until one of these happens.</span></span> <span data-ttu-id="7541d-160">Os exemplos a seguir mostram essas duas abordagens.</span><span class="sxs-lookup"><span data-stu-id="7541d-160">The following examples show these two approaches.</span></span>
>
>```js    
>Office.onReady();
>```
>
>
>```js
>Office.initialize = function () {};
>```

## <a name="see-also"></a><span data-ttu-id="7541d-161">Confira também</span><span class="sxs-lookup"><span data-stu-id="7541d-161">See also</span></span>

- [<span data-ttu-id="7541d-162">Entendendo a API JavaScript do Office</span><span class="sxs-lookup"><span data-stu-id="7541d-162">Understanding the Office JavaScript API</span></span>](understanding-the-javascript-api-for-office.md)
- [<span data-ttu-id="7541d-163">Carregando o DOM e o ambiente de tempo de execução</span><span class="sxs-lookup"><span data-stu-id="7541d-163">Loading the DOM and runtime environment</span></span>](loading-the-dom-and-runtime-environment.md)