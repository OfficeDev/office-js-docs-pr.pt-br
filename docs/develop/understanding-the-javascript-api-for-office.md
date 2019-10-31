---
title: Noções básicas da API JavaScript para Office
description: ''
ms.date: 06/21/2019
localization_priority: Priority
ms.openlocfilehash: 1954457b477472b8940841bb1ffe5954e49e01ec
ms.sourcegitcommit: b3996b1444e520b44cf752e76eef50908386ca26
ms.translationtype: HT
ms.contentlocale: pt-BR
ms.lasthandoff: 10/16/2019
ms.locfileid: "37524231"
---
# <a name="understanding-the-javascript-api-for-office"></a><span data-ttu-id="a60a2-102">Noções básicas da API JavaScript para Office</span><span class="sxs-lookup"><span data-stu-id="a60a2-102">Understanding the JavaScript API for Office</span></span>

<span data-ttu-id="a60a2-p101">Este artigo fornece informações sobre a API JavaScript para Office e como usá-la. Para referenciar as informações, consulte [API JavaScript para Office](/office/dev/add-ins/reference/javascript-api-for-office). Para obter informações sobre como atualizar os arquivos de projeto do Visual Studio para a versão mais recente da API JavaScript para Office, consulte [Atualizar a versão da API JavaScript para Office e arquivos de esquema do manifesto](update-your-javascript-api-for-office-and-manifest-schema-version.md).</span><span class="sxs-lookup"><span data-stu-id="a60a2-p101">This article provides information about the JavaScript API for Office and how to use it. For reference information, see [JavaScript API for Office](/office/dev/add-ins/reference/javascript-api-for-office). For information about updating Visual Studio project files to the most current version of the JavaScript API for Office, see [Update the version of your JavaScript API for Office and manifest schema files](update-your-javascript-api-for-office-and-manifest-schema-version.md).</span></span>

> [!NOTE]
> <span data-ttu-id="a60a2-p102">Caso pretenda [publicar](../publish/publish.md) o suplemento na experiência do Office depois de criá-lo, verifique se você está em conformidade com as [Políticas de validação do AppSource](/office/dev/store/validation-policies). Por exemplo, para passar na validação, seu suplemento deve funcionar em todas as plataformas com suporte aos métodos que você definir (para mais informações, confira a [seção 4.12](/office/dev/store/validation-policies#4-apps-and-add-ins-behave-predictably) e a [Página de hospedagem e disponibilidade de suplementos do Office](../overview/office-add-in-availability.md)).</span><span class="sxs-lookup"><span data-stu-id="a60a2-p102">If you plan to [publish](../publish/publish.md) your add-in to AppSource and make it available within the Office experience, make sure that you conform to the [AppSource validation policies](/office/dev/store/validation-policies). For example, to pass validation, your add-in must work across all platforms that support the methods that you define (for more information, see [section 4.12](/office/dev/store/validation-policies#4-apps-and-add-ins-behave-predictably) and the [Office Add-in host and availability page](../overview/office-add-in-availability.md)).</span></span> 

## <a name="referencing-the-javascript-api-for-office-library-in-your-add-in"></a><span data-ttu-id="a60a2-108">Fazer referência à biblioteca da API JavaScript para Office no suplemento</span><span class="sxs-lookup"><span data-stu-id="a60a2-108">Referencing the JavaScript API for Office library in your add-in</span></span>

<span data-ttu-id="a60a2-p103">A biblioteca da [API JavaScript para Office](/office/dev/add-ins/reference/javascript-api-for-office) consiste no arquivo Office.js e nos arquivos .js específicos do aplicativo de host associado, como Excel-15.js e Outlook-15.js. O método mais simples de fazer referência à API é usando nossa CDN e adicionando o seguinte `<script>` à marca `<head>` da sua página:</span><span class="sxs-lookup"><span data-stu-id="a60a2-p103">The [JavaScript API for Office](/office/dev/add-ins/reference/javascript-api-for-office) library consists of the Office.js file and associated host application-specific .js files, such as Excel-15.js and Outlook-15.js. The simplest method of referencing the API is using our CDN by adding the following `<script>` to your page's `<head>` tag:</span></span>  

```html
<script src="https://appsforoffice.microsoft.com/lib/1/hosted/office.js" type="text/javascript"></script>
```

<span data-ttu-id="a60a2-111">Isso baixará e colocará os arquivos da API JavaScript para Office em cache quando o suplemento for carregado pela primeira vez a fim de garantir que o suplemento esteja usando a implementação mais recente do Office.js e de seus arquivos associados na versão especificada.</span><span class="sxs-lookup"><span data-stu-id="a60a2-111">This will download and cache the JavaScript API for Office files the first time your add-in loads to make sure that it is using the most up-to-date implementation of Office.js and its associated files for the specified version.</span></span>

<span data-ttu-id="a60a2-112">Para saber mais sobre a CDN do Office.js, inclusive como é feito o controle de versão e como lidar com a compatibilidade com versões anteriores, confira [Fazendo referência à biblioteca da API JavaScript para Office na CDN (rede de distribuição de conteúdo)](referencing-the-javascript-api-for-office-library-from-its-cdn.md).</span><span class="sxs-lookup"><span data-stu-id="a60a2-112">For more details around the Office.js CDN, including how versioning and backward compatibility is handled, see [Referencing the JavaScript API for Office library from its content delivery network (CDN)](referencing-the-javascript-api-for-office-library-from-its-cdn.md).</span></span>

## <a name="initializing-your-add-in"></a><span data-ttu-id="a60a2-113">Inicialização do suplemento</span><span class="sxs-lookup"><span data-stu-id="a60a2-113">Initializing your add-in</span></span>

<span data-ttu-id="a60a2-114">**Aplica-se a:** todos os tipos de suplementos</span><span class="sxs-lookup"><span data-stu-id="a60a2-114">**Applies to:** All add-in types</span></span>

<span data-ttu-id="a60a2-115">Os Suplementos do Office têm sempre uma lógica de inicialização para fazer coisas como:</span><span class="sxs-lookup"><span data-stu-id="a60a2-115">Office Add-ins often have start-up logic to do things such as:</span></span>

- <span data-ttu-id="a60a2-116">Verificar se a versão do Office do usuário será compatível com todas as APIs do Office chamadas pelo código.</span><span class="sxs-lookup"><span data-stu-id="a60a2-116">Check that the user's version of Office will support all the Office APIs that your code calls.</span></span>

- <span data-ttu-id="a60a2-117">Garantir a existência de determinados artefatos, como uma planilha de nome específico.</span><span class="sxs-lookup"><span data-stu-id="a60a2-117">Ensure the existence of certain artifacts, such as worksheet with a specific name.</span></span>

- <span data-ttu-id="a60a2-118">Solicitar ao usuário selecionar algumas células no Excel e inserir um gráfico inicializado com esses valores selecionados.</span><span class="sxs-lookup"><span data-stu-id="a60a2-118">Prompting the user to select some cells in Excel, and then inserting a chart initialized with those selected values.</span></span>

- <span data-ttu-id="a60a2-119">Estabeleça associações.</span><span class="sxs-lookup"><span data-stu-id="a60a2-119">Establish bindings.</span></span>

- <span data-ttu-id="a60a2-120">Usar a API de caixa de diálogo do Office para solicitar ao usuário definir valores de configurações padrão para o suplemento.</span><span class="sxs-lookup"><span data-stu-id="a60a2-120">Use the Office dialog API to prompt the user for default add-in settings values.</span></span>

<span data-ttu-id="a60a2-121">O código de inicialização só deverá chamar as APIs do Office.js quando a biblioteca estiver carregada.</span><span class="sxs-lookup"><span data-stu-id="a60a2-121">But your start-up code must not call any Office.js APIs until the library is loaded.</span></span> <span data-ttu-id="a60a2-122">Há duas maneiras pelas quais o código pode garantir que a biblioteca seja carregada.</span><span class="sxs-lookup"><span data-stu-id="a60a2-122">There are two ways that your code can ensure that the library is loaded.</span></span> <span data-ttu-id="a60a2-123">Elas estão descritas nas seções a seguir:</span><span class="sxs-lookup"><span data-stu-id="a60a2-123">They are described in the following sections:</span></span> 

- [<span data-ttu-id="a60a2-124">Inicializar com Office.onReady()</span><span class="sxs-lookup"><span data-stu-id="a60a2-124">Initialize with Office.onReady()</span></span>](#initialize-with-officeonready)
- [<span data-ttu-id="a60a2-125">Inicializar com Office.initialize</span><span class="sxs-lookup"><span data-stu-id="a60a2-125">Initialize with Office.initialize</span></span>](#initialize-with-officeinitialize)

> [!TIP]
> <span data-ttu-id="a60a2-126">Recomendamos que use `Office.onReady()`em vez de`Office.initialize`.</span><span class="sxs-lookup"><span data-stu-id="a60a2-126">We recommend that you use `Office.onReady()` instead of `Office.initialize`.</span></span> <span data-ttu-id="a60a2-127">Embora `Office.initialize` ainda tenha suporte, usar`Office.onReady()` oferece mais flexibilidade.</span><span class="sxs-lookup"><span data-stu-id="a60a2-127">Although `Office.initialize` is still supported, using `Office.onReady()` provides more flexibility.</span></span> <span data-ttu-id="a60a2-128">É possível atribuir apenas um manipulador a `Office.initialize`, e ela é chamada apenas uma vez pela infraestrutura do Office. Mas você pode chamar `Office.onReady()` em diferentes locais no código, e usar diferentes retornos de chamadas.</span><span class="sxs-lookup"><span data-stu-id="a60a2-128">You can assign only one handler to `Office.initialize` and it's called only once by the Office infrastructure, but you can call `Office.onReady()` in different places in your code and use different callbacks.</span></span>
> 
> <span data-ttu-id="a60a2-129">Para saber mais sobre as diferenças entre essas técnicas, veja [Principais diferenças entre Office.initialize e Office.onReady()](#major-differences-between-officeinitialize-and-officeonready).</span><span class="sxs-lookup"><span data-stu-id="a60a2-129">For information about the differences in these techniques, see [Major differences between Office.initialize and Office.onReady()](#major-differences-between-officeinitialize-and-officeonready).</span></span>

<span data-ttu-id="a60a2-130">Para saber mais sobre a sequência de eventos na inicialização do suplemento, confira [Carregar o ambiente de tempo de execução e o DOM](loading-the-dom-and-runtime-environment.md).</span><span class="sxs-lookup"><span data-stu-id="a60a2-130">For more details about the sequence of events when an add-in is initialized, see [Loading the DOM and runtime environment](loading-the-dom-and-runtime-environment.md).</span></span>

### <a name="initialize-with-officeonready"></a><span data-ttu-id="a60a2-131">Inicializar com o Office.onReady()</span><span class="sxs-lookup"><span data-stu-id="a60a2-131">Initialize with Office.onReady()</span></span>

<span data-ttu-id="a60a2-132">`Office.onReady()` é um método assíncrono que retorna um objeto Promise enquanto verifica se a biblioteca do Office.js está carregada.</span><span class="sxs-lookup"><span data-stu-id="a60a2-132">`Office.onReady()` is an asynchronous method that returns a Promise object while it checks to see if the Office.js library is loaded.</span></span> <span data-ttu-id="a60a2-133">Somente quando a biblioteca é carregada, ela resolve o Promise como um objeto que especifica o aplicativo host do Office com um valor de enumeração `Office.HostType` (`Excel`, `Word` etc.), e a plataforma com um valor de enumeração `Office.PlatformType` (`PC`, `Mac`, `OfficeOnline` etc.).</span><span class="sxs-lookup"><span data-stu-id="a60a2-133">When the library is loaded, it resolves the Promise as an object that specifies the Office host application with an `Office.HostType` enum value (`Excel`, `Word`, etc.) and the platform with an `Office.PlatformType` enum value (`PC`, `Mac`, `OfficeOnline`, etc.).</span></span> <span data-ttu-id="a60a2-134">O Promise será resolvido imediatamente quando a biblioteca estiver carregada ao`Office.onReady()` ser chamada.</span><span class="sxs-lookup"><span data-stu-id="a60a2-134">The Promise resolves immediately if the library is already loaded when `Office.onReady()` is called.</span></span>

<span data-ttu-id="a60a2-135">Uma maneira de chamar `Office.onReady()` é passá-la por um método de retorno de chamada.</span><span class="sxs-lookup"><span data-stu-id="a60a2-135">One way to call `Office.onReady()` is to pass it a callback method.</span></span> <span data-ttu-id="a60a2-136">Exemplo:</span><span class="sxs-lookup"><span data-stu-id="a60a2-136">Here's an example:</span></span>

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

<span data-ttu-id="a60a2-137">Como alternativa, é possível encadear um método `then()` à chamada de `Office.onReady()`, em vez de passar um retorno de chamada.</span><span class="sxs-lookup"><span data-stu-id="a60a2-137">Alternatively, you can chain a `then()` method to the call of `Office.onReady()`, instead of passing a callback.</span></span> <span data-ttu-id="a60a2-138">Por exemplo, o código a seguir verifica se a versão do Excel do usuário é compatível com todas as APIs que o suplemento pode chamar.</span><span class="sxs-lookup"><span data-stu-id="a60a2-138">For example, the following code checks to see that the user's version of Excel supports all the APIs that the add-in might call.</span></span>

```js
Office.onReady()
    .then(function() {
        if (!Office.context.requirements.isSetSupported('ExcelApi', '1.7')) {
            console.log("Sorry, this add-in only works with newer versions of Excel.");
        }
    });
```

<span data-ttu-id="a60a2-139">Este é o mesmo exemplo que usa as palavras-chave `async` e `await` em TypeScript:</span><span class="sxs-lookup"><span data-stu-id="a60a2-139">Here is the same example using the `async` and `await` keywords in TypeScript:</span></span>

```typescript
(async () => {
    await Office.onReady();
    if (!Office.context.requirements.isSetSupported('ExcelApi', '1.7')) {
        console.log("Sorry, this add-in only works with newer versions of Excel.");
    }
})();
```

<span data-ttu-id="a60a2-140">Se estiver usando estruturas JavaScript adicionais que incluam testes e manipuladores próprios de inicialização, *geralmente* eles devem ser colocados dentro da resposta para `Office.onReady()`.</span><span class="sxs-lookup"><span data-stu-id="a60a2-140">If you are using additional JavaScript frameworks that include their own initialization handler or tests, these should *usually* be placed within the response to `Office.onReady()`.</span></span> <span data-ttu-id="a60a2-141">Por exemplo, a função [JQuery](https://jquery.com) `$(document).ready()` seria referenciada da seguinte maneira:</span><span class="sxs-lookup"><span data-stu-id="a60a2-141">For example, [JQuery's](https://jquery.com) `$(document).ready()` function would be referenced as follows:</span></span>

```js
Office.onReady(function() {
    // Office is ready
    $(document).ready(function () {
        // The document is ready
    });
});
```

<span data-ttu-id="a60a2-142">No entanto, há exceções a essa prática.</span><span class="sxs-lookup"><span data-stu-id="a60a2-142">However, there are exceptions to this practice.</span></span> <span data-ttu-id="a60a2-143">Por exemplo, digamos que você queira abrir o suplemento em um navegador (em vez de fazer sideload em um host do Office) para depurar a interface do usuário com ferramentas de navegador.</span><span class="sxs-lookup"><span data-stu-id="a60a2-143">For example, suppose you want to open your add-in in a browser (instead of sideload it in an Office host) in order to debug your UI with browser tools.</span></span> <span data-ttu-id="a60a2-144">Já que o Office.js não será carregado no navegador, `onReady` não será executado e o `$(document).ready` não será executado quando chamado dentro de `onReady` no Office.</span><span class="sxs-lookup"><span data-stu-id="a60a2-144">Since Office.js won't load in the browser, `onReady` won't run and the `$(document).ready` won't run if it's called inside the Office `onReady`.</span></span> <span data-ttu-id="a60a2-145">Outra exceção: você deseja que um indicador de progresso seja exibido no painel de tarefas enquanto o suplemento está sendo carregado.</span><span class="sxs-lookup"><span data-stu-id="a60a2-145">Another exception: you want a progress indicator to appear in the task pane while the add-in is loading.</span></span> <span data-ttu-id="a60a2-146">Nesse cenário, o código deve chamar `ready` da jQuery e usa a respectiva chamada de retorno para renderizar o indicador de progresso.</span><span class="sxs-lookup"><span data-stu-id="a60a2-146">In this scenario, your code should call the jQuery `ready` and use it's callback to render the progress indicator.</span></span> <span data-ttu-id="a60a2-147">Em seguida, a chamada de retorno do Office `onReady` pode substituir o indicador de progresso com a interface do usuário final.</span><span class="sxs-lookup"><span data-stu-id="a60a2-147">Then the Office `onReady`'s callback can replace the progress indicator with the final UI.</span></span> 

### <a name="initialize-with-officeinitialize"></a><span data-ttu-id="a60a2-148">Inicializar com Office.initialize</span><span class="sxs-lookup"><span data-stu-id="a60a2-148">Initialize with Office.initialize</span></span>

<span data-ttu-id="a60a2-149">Um evento de inicialização é disparado quando a biblioteca do Office.js está carregada e pronta para a interação com o usuário.</span><span class="sxs-lookup"><span data-stu-id="a60a2-149">An initialize event fires when the Office.js library is loaded and ready for user interaction.</span></span> <span data-ttu-id="a60a2-150">É possível atribuir um manipulador ao `Office.initialize` que implementa a lógica de inicialização.</span><span class="sxs-lookup"><span data-stu-id="a60a2-150">You can assign a handler to `Office.initialize` that implements your initialization logic.</span></span> <span data-ttu-id="a60a2-151">Veja a seguir um exemplo que verifica se a versão do Excel do usuário é compatível com todas as APIs que o suplemento pode chamar.</span><span class="sxs-lookup"><span data-stu-id="a60a2-151">The following is an example that checks to see that the user's version of Excel supports all the APIs that the add-in might call.</span></span>

```js
Office.initialize = function () {
    if (!Office.context.requirements.isSetSupported('ExcelApi', '1.7')) {
        console.log("Sorry, this add-in only works with newer versions of Excel.");
    }
};
```

<span data-ttu-id="a60a2-152">Se estiver usando estruturas JavaScript adicionais que incluam testes e manipuladores próprios de inicialização, *geralmente* eles devem ser colocados dentro do evento `Office.initialize`.</span><span class="sxs-lookup"><span data-stu-id="a60a2-152">If you are using additional JavaScript frameworks that include their own initialization handler or tests, these should *usually* be placed within the `Office.initialize` event.</span></span> <span data-ttu-id="a60a2-153">No entanto, as exceções descritas anteriormente na seção **Inicializar com Office.onReady()** também se aplicam neste caso. Por exemplo, a função `$(document).ready()` do [JQuery](https://jquery.com) pode ser referenciada da seguinte maneira:</span><span class="sxs-lookup"><span data-stu-id="a60a2-153">(But the exceptions described in the **Initialize with Office.onReady()** section earlier apply in this case also.) For example, [JQuery's](https://jquery.com) `$(document).ready()` function would be referenced as follows:</span></span>

```js
Office.initialize = function () {
    // Office is ready
    $(document).ready(function () {
        // The document is ready
    });
  };
```

<span data-ttu-id="a60a2-154">Para suplementos de conteúdo e painel de tarefas, `Office.initialize` fornece um parâmetro _reason_ adicional.</span><span class="sxs-lookup"><span data-stu-id="a60a2-154">For task pane and content add-ins, `Office.initialize` provides an additional _reason_ parameter.</span></span> <span data-ttu-id="a60a2-155">Esse parâmetro especifica como um suplemento foi adicionado ao documento atual.</span><span class="sxs-lookup"><span data-stu-id="a60a2-155">This parameter specifies how an add-in was added to the current document.</span></span> <span data-ttu-id="a60a2-156">Você pode usar isso para fornecer uma lógica diferente para quando um suplemento é inserido pela primeira vez, em comparação com quando já existia dentro do documento.</span><span class="sxs-lookup"><span data-stu-id="a60a2-156">You can use this to provide different logic for when an add-in is first inserted versus when it already existed within the document.</span></span>

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

<span data-ttu-id="a60a2-157">Para saber mais, veja [Evento Office.initialize](/javascript/api/office) e [Enumeração da InitializationReason](/javascript/api/office/office.initializationreason).</span><span class="sxs-lookup"><span data-stu-id="a60a2-157">For more information, see [Office.initialize Event](/javascript/api/office) and [InitializationReason Enumeration](/javascript/api/office/office.initializationreason).</span></span>

### <a name="major-differences-between-officeinitialize-and-officeonready"></a><span data-ttu-id="a60a2-158">Principais diferenças entre Office.initialize e Office.onReady</span><span class="sxs-lookup"><span data-stu-id="a60a2-158">Major differences between Office.initialize and Office.onReady</span></span>

- <span data-ttu-id="a60a2-159">É possível atribuir apenas um manipulador a `Office.initialize`, e ela é chamada apenas uma vez pela infraestrutura do Office, mas você pode chamar `Office.onReady()` em diferentes locais no código, e usar diferentes retornos de chamadas.</span><span class="sxs-lookup"><span data-stu-id="a60a2-159">You can assign only one handler to `Office.initialize` and it's called only once by the Office infrastructure; but you can call `Office.onReady()` in different places in your code and use different callbacks.</span></span> <span data-ttu-id="a60a2-160">Por exemplo, o código pode chamar `Office.onReady()`, logo que o script personalizado é carregado com um retorno de chamada que executa uma lógica de inicialização. Além disso, o código pode ter um botão no painel de tarefas, cujo script chama `Office.onReady()` com um retorno de chamada diferente.</span><span class="sxs-lookup"><span data-stu-id="a60a2-160">For example, your code could call `Office.onReady()` as soon as your custom script loads with a callback that runs initialization logic; and your code could also have a button in the task pane, whose script calls `Office.onReady()` with a different callback.</span></span> <span data-ttu-id="a60a2-161">Quando isso ocorre, o segundo retorno de chamada é executado quando o botão é clicado.</span><span class="sxs-lookup"><span data-stu-id="a60a2-161">If so, the second callback runs when the button is clicked.</span></span>

- <span data-ttu-id="a60a2-162">O evento `Office.initialize` é disparado no final do processo interno, e que o Office.js é inicializado automaticamente.</span><span class="sxs-lookup"><span data-stu-id="a60a2-162">The `Office.initialize` event fires at the end of the internal process in which Office.js initializes itself.</span></span> <span data-ttu-id="a60a2-163">Ele também é disparado *imediatamente* após o término do processo interno.</span><span class="sxs-lookup"><span data-stu-id="a60a2-163">And it fires *immediately* after the internal process ends.</span></span> <span data-ttu-id="a60a2-164">Se o código no qual você atribui um manipulador ao evento for executado muito tempo após o evento ser disparado, então o manipulador não será executado.</span><span class="sxs-lookup"><span data-stu-id="a60a2-164">If the code in which you assign a handler to the event executes too long after the event fires, then your handler doesn't run.</span></span> <span data-ttu-id="a60a2-165">Por exemplo, se estiver usando o gerenciador de tarefas WebPack, ele poderá configurar a home page do suplemento para carregar arquivos de polyfill, após carregar o Office.js, mas antes de carregar o JavaScript personalizado.</span><span class="sxs-lookup"><span data-stu-id="a60a2-165">For example, if you are using the WebPack task manager, it might configure the add-in's home page to load polyfill files after it loads Office.js but before it loads your custom JavaScript.</span></span> <span data-ttu-id="a60a2-166">Quando o script carrega e atribui o manipulador, o evento de inicialização já ocorreu.</span><span class="sxs-lookup"><span data-stu-id="a60a2-166">By the time your script loads and assigns the handler, the initialize event has already happened.</span></span> <span data-ttu-id="a60a2-167">Mas nunca é "tarde demais" para chamar `Office.onReady()`.</span><span class="sxs-lookup"><span data-stu-id="a60a2-167">But it is never "too late" to call `Office.onReady()`.</span></span> <span data-ttu-id="a60a2-168">Caso o evento de inicialização já tenha ocorrido, o retorno de chamada é executado imediatamente.</span><span class="sxs-lookup"><span data-stu-id="a60a2-168">If the initialize event has already happened, the callback runs immediately.</span></span>

> [!NOTE]
> <span data-ttu-id="a60a2-169">Mesmo que não tenha uma lógica de inicialização, você deve atribuir ou chamar `Office.onReady()` uma função vazia para `Office.initialize` quando o JavaScript do suplemento for carregado.</span><span class="sxs-lookup"><span data-stu-id="a60a2-169">Even if you have no start-up logic, you should either call `Office.onReady()` or assign an empty function to `Office.initialize` when your add-in JavaScript loads.</span></span> <span data-ttu-id="a60a2-170">Algumas combinações de host e da plataforma do Office não carregam o painel de tarefas até uma das delas aconteça.</span><span class="sxs-lookup"><span data-stu-id="a60a2-170">Some Office host and platform combinations won't load the task pane until one of these happens.</span></span> <span data-ttu-id="a60a2-171">Os exemplos a seguir mostram essas duas abordagens.</span><span class="sxs-lookup"><span data-stu-id="a60a2-171">The following examples show these two approaches.</span></span>
>
>```js  
>Office.onReady();
>```
>
>
>```js
>Office.initialize = function () {};
>```

## <a name="office-javascript-api-object-model"></a><span data-ttu-id="a60a2-172">Modelo de objeto de API JavaScript para Office</span><span class="sxs-lookup"><span data-stu-id="a60a2-172">Office JavaScript API object model</span></span>

<span data-ttu-id="a60a2-173">Depois de inicializado, o suplemento pode interagir com o host (por exemplo, o Excel ou o Outlook).</span><span class="sxs-lookup"><span data-stu-id="a60a2-173">Once initialized, the add-in can interact with the host (e.g. Excel, Outlook).</span></span> <span data-ttu-id="a60a2-174">A página [Modelo de objeto de API JavaScript para Office](office-javascript-api-object-model.md) tem mais detalhes sobre padrões de uso específicos.</span><span class="sxs-lookup"><span data-stu-id="a60a2-174">The [Office JavaScript API object model](office-javascript-api-object-model.md) page has more details on specific usage patterns.</span></span> <span data-ttu-id="a60a2-175">Há também documentação de referência detalhada para [APIs Comuns](/office/dev/add-ins/reference/javascript-api-for-office) e hosts específicos.</span><span class="sxs-lookup"><span data-stu-id="a60a2-175">There is also detailed reference documentation for both [Common APIs](/office/dev/add-ins/reference/javascript-api-for-office) and host-specific APIs.</span></span>
