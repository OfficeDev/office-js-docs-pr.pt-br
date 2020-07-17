---
title: Usar a API da Caixa de Diálogo do Office nos suplementos do Office
description: Saiba mais sobre a criação de uma caixa de diálogo em um suplemento do Office.
ms.date: 06/10/2020
localization_priority: Normal
ms.openlocfilehash: 5cdd457b99636dd244eed1fa88c1b76cab23ee8c
ms.sourcegitcommit: 472b81642e9eb5fb2a55cd98a7b0826d37eb7f73
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 07/17/2020
ms.locfileid: "45159560"
---
# <a name="use-the-office-dialog-api-in-office-add-ins"></a><span data-ttu-id="3ff78-103">Usar a API de diálogo do Office em suplementos do Office</span><span class="sxs-lookup"><span data-stu-id="3ff78-103">Use the Office dialog API in Office Add-ins</span></span>

<span data-ttu-id="3ff78-104">Você pode usar a [API de Caixa de diálogo do Office](/javascript/api/office/office.ui) para abrir caixas de diálogo no seu Suplemento do Office.</span><span class="sxs-lookup"><span data-stu-id="3ff78-104">You can use the [Office dialog API](/javascript/api/office/office.ui) to open dialog boxes in your Office Add-in.</span></span> <span data-ttu-id="3ff78-105">Este artigo fornece orientações para usar a API de Caixa de diálogo em seu Suplemento do Office.</span><span class="sxs-lookup"><span data-stu-id="3ff78-105">This article provides guidance for using the dialog API in your Office Add-in.</span></span>

> [!NOTE]
> <span data-ttu-id="3ff78-p102">Para informações sobre os programas para os quais a API de Caixa de Diálogo tem suporte no momento, confira [Conjuntos de requisitos da API de Caixa de Diálogo](../reference/requirement-sets/dialog-api-requirement-sets.md). Atualmente, a API de Caixa de Diálogo tem suporte para Word, Excel, PowerPoint e Outlook.</span><span class="sxs-lookup"><span data-stu-id="3ff78-p102">For information about where the Dialog API is currently supported, see [Dialog API requirement sets](../reference/requirement-sets/dialog-api-requirement-sets.md). The Dialog API is currently supported for Word, Excel, PowerPoint, and Outlook.</span></span>

<span data-ttu-id="3ff78-108">Um cenário fundamental para a API de Caixa de Diálogo é habilitar a autenticação com um recurso como o Google, o Facebook ou o Microsoft Graph.</span><span class="sxs-lookup"><span data-stu-id="3ff78-108">A primary scenario for the Dialog API is to enable authentication with a resource such as Google, Facebook, or Microsoft Graph.</span></span> <span data-ttu-id="3ff78-109">Para saber mais, confira [ autenticação com APIs de Caixa de Diálogo do Office](auth-with-office-dialog-api.md) *depois* que você se familiarizar com este artigo.</span><span class="sxs-lookup"><span data-stu-id="3ff78-109">For more information, see [Authenticate with the Office dialog API](auth-with-office-dialog-api.md) *after* you are familiar with this article.</span></span>

<span data-ttu-id="3ff78-110">Considere abrir uma caixa de diálogo em um painel de tarefas, suplemento de conteúdo ou [comando de suplemento](../design/add-in-commands.md) para fazer o seguinte:</span><span class="sxs-lookup"><span data-stu-id="3ff78-110">Consider opening a dialog box from a task pane or content add-in or [add-in command](../design/add-in-commands.md) to do the following:</span></span>

- <span data-ttu-id="3ff78-111">Exibir páginas de entrada que não podem ser abertas diretamente em um painel de tarefas.</span><span class="sxs-lookup"><span data-stu-id="3ff78-111">Display sign in pages that cannot be opened directly in a task pane.</span></span>
- <span data-ttu-id="3ff78-112">Fornecer mais espaço na tela, ou até uma tela inteira, para algumas tarefas no seu suplemento.</span><span class="sxs-lookup"><span data-stu-id="3ff78-112">Provide more screen space, or even a full screen, for some tasks in your add-in.</span></span>
- <span data-ttu-id="3ff78-113">Hospedar um vídeo que seria muito pequeno se fosse confinado em um painel de tarefas.</span><span class="sxs-lookup"><span data-stu-id="3ff78-113">Host a video that would be too small if confined to a task pane.</span></span>

> [!NOTE]
> <span data-ttu-id="3ff78-114">Como a sobreposição de elementos de IU não são recomendáveis, evite abrir uma caixa de diálogo em um painel de tarefas a menos que seu cenário o obrigue a fazer isso.</span><span class="sxs-lookup"><span data-stu-id="3ff78-114">Because overlapping UI elements are discouraged, avoid opening a dialog box from a task pane unless your scenario requires it.</span></span> <span data-ttu-id="3ff78-115">Ao considerar como usar a área de superfície de um painel de tarefas, observe que painéis de tarefas podem ter guias.</span><span class="sxs-lookup"><span data-stu-id="3ff78-115">When you consider how to use the surface area of a task pane, note that task panes can be tabbed.</span></span> <span data-ttu-id="3ff78-116">Por exemplo, confira o exemplo [Suplemento do Excel JavaScriptSalesTracker](https://github.com/OfficeDev/Excel-Add-in-JavaScript-SalesTracker).</span><span class="sxs-lookup"><span data-stu-id="3ff78-116">For an example, see the [Excel Add-in JavaScript SalesTracker](https://github.com/OfficeDev/Excel-Add-in-JavaScript-SalesTracker) sample.</span></span>

<span data-ttu-id="3ff78-117">A imagem abaixo mostra um exemplo de uma caixa de diálogo.</span><span class="sxs-lookup"><span data-stu-id="3ff78-117">The following image shows an example of a dialog box.</span></span>

![Comandos de suplemento](../images/auth-o-dialog-open.png)

<span data-ttu-id="3ff78-119">A caixa de diálogo sempre abre no centro da tela.</span><span class="sxs-lookup"><span data-stu-id="3ff78-119">Note that the dialog box always opens in the center of the screen.</span></span> <span data-ttu-id="3ff78-120">O usuário pode movê-la e redimensioná-la.</span><span class="sxs-lookup"><span data-stu-id="3ff78-120">The user can move and resize it.</span></span> <span data-ttu-id="3ff78-121">A janela é *não modal*: o usuário pode continuar a interagir com o documento no aplicativo do Office do host e com a página no painel de tarefas, caso houver uma.</span><span class="sxs-lookup"><span data-stu-id="3ff78-121">The window is *nonmodal*--a user can continue to interact with both the document in the host Office application and with the page in the task pane, if there is one.</span></span>

## <a name="open-a-dialog-box-from-a-host-page"></a><span data-ttu-id="3ff78-122">Abrir uma caixa de diálogo em uma página de host</span><span class="sxs-lookup"><span data-stu-id="3ff78-122">Open a dialog box from a host page</span></span>

<span data-ttu-id="3ff78-123">As APIs JavaScript para Office incluem um objeto[Dialog](/javascript/api/office/office.dialog) e duas funções no [namespace Office.context.ui](/javascript/api/office/office.ui).</span><span class="sxs-lookup"><span data-stu-id="3ff78-123">The Office JavaScript APIs include a [Dialog](/javascript/api/office/office.dialog) object and two functions in the [Office.context.ui namespace](/javascript/api/office/office.ui).</span></span>

<span data-ttu-id="3ff78-124">Para abrir uma caixa de diálogo, seu código, geralmente uma página no painel de tarefas chama o método [displayDialogAsync](/javascript/api/office/office.ui) e transmite a ele a URL do recurso que você deseja abrir.</span><span class="sxs-lookup"><span data-stu-id="3ff78-124">To open a dialog box, your code, typically a page in a task pane, calls the [displayDialogAsync](/javascript/api/office/office.ui) method and passes to it the URL of the resource that you want to open.</span></span> <span data-ttu-id="3ff78-125">A página em que esse método é chamado é conhecida como "página host".</span><span class="sxs-lookup"><span data-stu-id="3ff78-125">The page on which this method is called is known as the "host page".</span></span> <span data-ttu-id="3ff78-126">Por exemplo, se você chamar esse método no script index.html em um painel de tarefas, index.html será a página do host da caixa de diálogo que o método abre.</span><span class="sxs-lookup"><span data-stu-id="3ff78-126">For example, if you call this method in script on index.html in a task pane, then index.html is the host page of the dialog box that the method opens.</span></span>

<span data-ttu-id="3ff78-127">O recurso aberto na página de diálogo geralmente é uma página, mas pode ser um método controlador em um aplicativo MVC, uma rota, um método de serviço Web ou qualquer outro recurso.</span><span class="sxs-lookup"><span data-stu-id="3ff78-127">The resource that is opened in the dialog box is usually a page, but it can be a controller method in an MVC application, a route, a web service method, or any other resource.</span></span> <span data-ttu-id="3ff78-128">Neste artigo, 'página' ou 'site' refere-se ao recurso na caixa de diálogo.</span><span class="sxs-lookup"><span data-stu-id="3ff78-128">In this article, 'page' or 'website' refers to the resource in the dialog box.</span></span> <span data-ttu-id="3ff78-129">O código a seguir é um exemplo simples:</span><span class="sxs-lookup"><span data-stu-id="3ff78-129">The following code is a simple example:</span></span>

```js
Office.context.ui.displayDialogAsync('https://myAddinDomain/myDialog.html');
```

> [!NOTE]
> - <span data-ttu-id="3ff78-130">A URL usa o protocolo HTTP**S**.</span><span class="sxs-lookup"><span data-stu-id="3ff78-130">The URL uses the HTTP**S** protocol.</span></span> <span data-ttu-id="3ff78-131">Isso é obrigatório para todas as páginas carregadas em uma caixa diálogo, não apenas para a primeira página carregada.</span><span class="sxs-lookup"><span data-stu-id="3ff78-131">This is mandatory for all pages loaded in a dialog box, not just the first page loaded.</span></span>
> - <span data-ttu-id="3ff78-132">A caixa de diálogo é igual ao domínio da página de host, que pode ser a página em um painel de tarefas ou o [arquivo de função](../reference/manifest/functionfile.md) de um comando de suplemento.</span><span class="sxs-lookup"><span data-stu-id="3ff78-132">The dialog box's domain is the same as the domain of the host page, which can be the page in a task pane or the [function file](../reference/manifest/functionfile.md) of an add-in command.</span></span> <span data-ttu-id="3ff78-133">Isso é necessário: a página, o método do controlador ou outro recurso que é passado para o método `displayDialogAsync` deve estar no mesmo domínio que a página de host.</span><span class="sxs-lookup"><span data-stu-id="3ff78-133">This is required: the page, controller method, or other resource that is passed to the `displayDialogAsync` method must be in the same domain as the host page.</span></span>

> [!IMPORTANT]
> <span data-ttu-id="3ff78-134">A página de host e o recurso que abrem na caixa de diálogo devem ter o mesmo domínio inteiro.</span><span class="sxs-lookup"><span data-stu-id="3ff78-134">The host page and the resource that opens in the dialog box must have the same full domain.</span></span> <span data-ttu-id="3ff78-135">Se você tentar passar `displayDialogAsync` para um subdomínio do domínio do suplemento, ele não funcionará.</span><span class="sxs-lookup"><span data-stu-id="3ff78-135">If you attempt to pass `displayDialogAsync` a subdomain of the add-in's domain, it will not work.</span></span> <span data-ttu-id="3ff78-136">O domínio completo, incluindo qualquer subdomínio, deve corresponder.</span><span class="sxs-lookup"><span data-stu-id="3ff78-136">The full domain, including any subdomain, must match.</span></span>

<span data-ttu-id="3ff78-137">Após o carregamento da primeira página (ou de outro recurso), um usuário pode usar links ou outra interface de usuário para qualquer site (ou outro recurso) que usa HTTPS.</span><span class="sxs-lookup"><span data-stu-id="3ff78-137">After the first page (or other resource) is loaded, a user can use links or other UI to navigate to any website (or other resource) that uses HTTPS.</span></span> <span data-ttu-id="3ff78-138">Também é possível criar a primeira página para redirecionar imediatamente para outro site.</span><span class="sxs-lookup"><span data-stu-id="3ff78-138">You can also design the first page to immediately redirect to another site.</span></span>

<span data-ttu-id="3ff78-139">Por padrão, a caixa de diálogo ocupará 80% da altura e da largura na tela do dispositivo, mas você pode definir porcentagens diferentes. Basta transmitir um objeto de configuração para o método, como mostra o exemplo a seguir.</span><span class="sxs-lookup"><span data-stu-id="3ff78-139">By default, the dialog box will occupy 80% of the height and width of the device screen, but you can set different percentages by passing a configuration object to the method, as shown in the following example:</span></span>

```js
Office.context.ui.displayDialogAsync('https://myDomain/myDialog.html', {height: 30, width: 20});
```

<span data-ttu-id="3ff78-140">Para ver um suplemento de exemplo que faz isso, confira [Exemplo de API de Caixa de diálogo do Suplemento do Office](https://github.com/OfficeDev/Office-Add-in-Dialog-API-Simple-Example).</span><span class="sxs-lookup"><span data-stu-id="3ff78-140">For a sample add-in that does this, see [Office Add-in Dialog API Example](https://github.com/OfficeDev/Office-Add-in-Dialog-API-Simple-Example).</span></span>

<span data-ttu-id="3ff78-p112">Defina os dois valores como 100% para ter uma verdadeira experiência de tela inteira. O máximo real é 99,5%, e a janela ainda poderá ser movida e redimensionada.</span><span class="sxs-lookup"><span data-stu-id="3ff78-p112">Set both values to 100% to get what is effectively a full screen experience. (The effective maximum is 99.5%, and the window is still moveable and resizable.)</span></span>

> [!NOTE]
> <span data-ttu-id="3ff78-p113">Apenas uma caixa de diálogo pode ser aberta em uma janela do host. Tentar abrir outra caixa de diálogo gera um erro. Portanto, por exemplo, se um usuário abrir uma caixa de diálogo no painel de tarefas, ele não poderá abrir uma segunda caixa de diálogo em uma página diferente no painel de tarefas. No entanto, quando uma caixa de diálogo é aberta em um [comando de suplemento](../design/add-in-commands.md), o comando abre um arquivo HTML novo (mas não visto) sempre que ele é selecionado. Isso cria uma nova janela do host (não vista) para que cada janela possa iniciar sua própria caixa de diálogo. Para obter mais informações, confira [Erros de displayDialogAsync](dialog-handle-errors-events.md#errors-from-displaydialogasync).</span><span class="sxs-lookup"><span data-stu-id="3ff78-p113">You can open only one dialog box from a host window. An attempt to open another dialog box generates an error. For example, if a user opens a dialog box from a task pane, she cannot open a second dialog box, from a different page in the task pane. However, when a dialog box is opened from an [add-in command](../design/add-in-commands.md), the command opens a new (but unseen) HTML file each time it is selected. This creates a new (unseen) host window, so each such window can launch its own dialog box. For more information, see [Errors from displayDialogAsync](dialog-handle-errors-events.md#errors-from-displaydialogasync).</span></span>

### <a name="take-advantage-of-a-performance-option-in-office-on-the-web"></a><span data-ttu-id="3ff78-149">Aproveite uma opção de desempenho no Office na Web</span><span class="sxs-lookup"><span data-stu-id="3ff78-149">Take advantage of a performance option in Office on the web</span></span>

<span data-ttu-id="3ff78-150">A propriedade `displayInIframe` é uma propriedade adicional no objeto de configuração que você pode passar para o`displayDialogAsync`.</span><span class="sxs-lookup"><span data-stu-id="3ff78-150">The `displayInIframe` property is an additional property in the configuration object that you can pass to `displayDialogAsync`.</span></span> <span data-ttu-id="3ff78-151">Quando essa propriedade for definida como `true` e o suplemento estiver em execução em um documento aberto no Office Online, a caixa de diálogo será aberta como um iframe flutuante, em vez de uma janela independente, o que faz com que ela seja aberta mais rapidamente.</span><span class="sxs-lookup"><span data-stu-id="3ff78-151">When this property is set to `true`, and the add-in is running in a document opened in Office on the web, the dialog box will open as a floating iframe rather than an independent window, which makes it open faster.</span></span> <span data-ttu-id="3ff78-152">Este é um exemplo:</span><span class="sxs-lookup"><span data-stu-id="3ff78-152">The following is an example:</span></span>

```js
Office.context.ui.displayDialogAsync('https://myDomain/myDialog.html', {height: 30, width: 20, displayInIframe: true});
```

<span data-ttu-id="3ff78-153">O valor padrão é `false`, que é o mesmo que omitir a propriedade inteiramente.</span><span class="sxs-lookup"><span data-stu-id="3ff78-153">The default value is `false`, which is the same as omitting the property entirely.</span></span> <span data-ttu-id="3ff78-154">Se o suplemento não estiver sendo executado no Office Online, o `displayInIframe` será ignorado.</span><span class="sxs-lookup"><span data-stu-id="3ff78-154">If the add-in is not running in Office on the web, the `displayInIframe` is ignored.</span></span>

> [!NOTE]
> <span data-ttu-id="3ff78-155">Você **não** deverá usar `displayInIframe: true` se a caixa de diálogo redirecionar a qualquer ponto para uma página que não possa ser aberta em um iframe.</span><span class="sxs-lookup"><span data-stu-id="3ff78-155">You should **not** use `displayInIframe: true` if the dialog box will at any point redirect to a page that cannot be opened in an iframe.</span></span> <span data-ttu-id="3ff78-156">Por exemplo, as páginas de entrada de muitos serviços Web populares, como a conta do Google e da Microsoft, não podem ser abertas em um iframe.</span><span class="sxs-lookup"><span data-stu-id="3ff78-156">For example, the sign in pages of many popular web services, such as Google and Microsoft account, cannot be opened in an iframe.</span></span>

## <a name="send-information-from-the-dialog-box-to-the-host-page"></a><span data-ttu-id="3ff78-157">Envie informações da caixa de diálogo para a página host</span><span class="sxs-lookup"><span data-stu-id="3ff78-157">Send information from the dialog box to the host page</span></span>

<span data-ttu-id="3ff78-158">A caixa de diálogo não pode se comunicar com a página host no painel de tarefas, a menos que:</span><span class="sxs-lookup"><span data-stu-id="3ff78-158">The dialog box cannot communicate with the host page in the task pane unless:</span></span>

- <span data-ttu-id="3ff78-159">A página atual na caixa de diálogo esteja no mesmo domínio da página host.</span><span class="sxs-lookup"><span data-stu-id="3ff78-159">The current page in the dialog box is in the same domain as the host page.</span></span>
- <span data-ttu-id="3ff78-p117">A biblioteca da API JavaScript do Office é carregada na página. (Como qualquer página que usa a biblioteca da API JavaScript do Office, o script para a página deve atribuir um método à `Office.initialize` propriedade, embora possa ser um método vazio. Para obter detalhes, consulte [inicializar o suplemento do Office](initialize-add-in.md).</span><span class="sxs-lookup"><span data-stu-id="3ff78-p117">The Office JavaScript API library is loaded in the page. (Like any page that uses the Office JavaScript API library, script for the page must assign a method to the `Office.initialize` property, although it can be an empty method. For details, see [Initialize your Office Add-in](initialize-add-in.md).)</span></span>

<span data-ttu-id="3ff78-163">O código na caixa de diálogo use a função [messageParent](/javascript/api/office/office.ui#messageparent-message-) para enviar uma mensagem de cadeia de caracteres ou um valor booliano para a página host.</span><span class="sxs-lookup"><span data-stu-id="3ff78-163">Code in the dialog box uses the [messageParent](/javascript/api/office/office.ui#messageparent-message-) function to send either a Boolean value or a string message to the host page.</span></span> <span data-ttu-id="3ff78-164">A cadeia de caracteres pode ser uma palavra, uma frase, um blob XML, um JSON em formato de cadeia de caracteres ou qualquer outra coisa que possa ser serializada em uma cadeia de caracteres.</span><span class="sxs-lookup"><span data-stu-id="3ff78-164">The string can be a word, sentence, XML blob, stringified JSON, or anything else that can be serialized to a string.</span></span> <span data-ttu-id="3ff78-165">Este é um exemplo:</span><span class="sxs-lookup"><span data-stu-id="3ff78-165">The following is an example:</span></span>

```js
if (loginSuccess) {
    Office.context.ui.messageParent(true);
}
```

> [!IMPORTANT]
> - <span data-ttu-id="3ff78-166">A função `messageParent` só pode ser chamada em uma página com o mesmo domínio (incluindo o protocolo e a porta) da página host.</span><span class="sxs-lookup"><span data-stu-id="3ff78-166">The `messageParent` function can only be called on a page with the same domain (including protocol and port) as the host page.</span></span>
> - <span data-ttu-id="3ff78-167">A `messageParent` função é uma das *only* duas APIs do Office js que podem ser chamadas na caixa de diálogo.</span><span class="sxs-lookup"><span data-stu-id="3ff78-167">The `messageParent` function is one of *only* two Office JS APIs that can be called in the dialog box.</span></span> 
> - <span data-ttu-id="3ff78-168">A outra API JS que pode ser chamada na caixa de diálogo é `Office.context.requirements.isSetSupported` .</span><span class="sxs-lookup"><span data-stu-id="3ff78-168">The other JS API that can be called in the dialog box is `Office.context.requirements.isSetSupported`.</span></span> <span data-ttu-id="3ff78-169">Para saber mais, confira [Especificar hosts do Office e requisitos da API](specify-office-hosts-and-api-requirements.md).</span><span class="sxs-lookup"><span data-stu-id="3ff78-169">For information about it, see [Specify Office hosts and API requirements](specify-office-hosts-and-api-requirements.md).</span></span> <span data-ttu-id="3ff78-170">No entanto, na caixa de diálogo, essa API não tem suporte no Outlook 2016 1-time Purchase (ou seja, a versão MSI).</span><span class="sxs-lookup"><span data-stu-id="3ff78-170">However, in the dialog box, this API isn't supported in Outlook 2016 one-time purchase (that is, the MSI version).</span></span>

<span data-ttu-id="3ff78-171">No próximo exemplo, `googleProfile` é uma versão em formato de cadeia de caracteres do perfil do Google do usuário.</span><span class="sxs-lookup"><span data-stu-id="3ff78-171">In the next example, `googleProfile` is a stringified version of the user's Google profile.</span></span>

```js
if (loginSuccess) {
    Office.context.ui.messageParent(googleProfile);
}
```

<span data-ttu-id="3ff78-p120">A página host deve ser configurada para receber a mensagem. Você pode fazer isso adicionando um parâmetro de retorno de chamada à chamada original de `displayDialogAsync`. O retorno de chamada atribui um manipulador ao evento `DialogMessageReceived`. Apresentamos um exemplo a seguir:</span><span class="sxs-lookup"><span data-stu-id="3ff78-p120">The host page must be configured to receive the message. You do this by adding a callback parameter to the original call of `displayDialogAsync`. The callback assigns a handler to the `DialogMessageReceived` event. The following is an example:</span></span>

```js
var dialog;
Office.context.ui.displayDialogAsync('https://myDomain/myDialog.html', {height: 30, width: 20},
    function (asyncResult) {
        dialog = asyncResult.value;
        dialog.addEventHandler(Office.EventType.DialogMessageReceived, processMessage);
    }
);
```

> [!NOTE]
> - <span data-ttu-id="3ff78-176">O Office transmite um objeto [AsyncResult](/javascript/api/office/office.asyncresult) para o retorno de chamada.</span><span class="sxs-lookup"><span data-stu-id="3ff78-176">Office passes an [AsyncResult](/javascript/api/office/office.asyncresult) object to the callback.</span></span> <span data-ttu-id="3ff78-177">Ele representa o resultado de tentativas de abrir a caixa de diálogo, </span><span class="sxs-lookup"><span data-stu-id="3ff78-177">It represents the result of the attempt to open the dialog box.</span></span> <span data-ttu-id="3ff78-178">Ela não representa o resultado de eventos na caixa diálogo.</span><span class="sxs-lookup"><span data-stu-id="3ff78-178">It does not represent the outcome of any events in the dialog box.</span></span> <span data-ttu-id="3ff78-179">Para saber mais sobre essa distinção, confira [Manipular erros e eventos](dialog-handle-errors-events.md).</span><span class="sxs-lookup"><span data-stu-id="3ff78-179">For more on this distinction, see [Handle errors and events](dialog-handle-errors-events.md).</span></span>
> - <span data-ttu-id="3ff78-180">A propriedade `value` do `asyncResult` é definida como um objeto [Dialog](/javascript/api/office/office.dialog) que existe na página host, não no contexto da execução da caixa de diálogo.</span><span class="sxs-lookup"><span data-stu-id="3ff78-180">The `value` property of the `asyncResult` is set to a [Dialog](/javascript/api/office/office.dialog) object, which exists in the host page, not in the dialog box's execution context.</span></span>
> - <span data-ttu-id="3ff78-p122">O `processMessage` é a função que manipula o evento. Você pode dar a ele o nome que desejar.</span><span class="sxs-lookup"><span data-stu-id="3ff78-p122">The `processMessage` is the function that handles the event. You can give it any name you want.</span></span>
> - <span data-ttu-id="3ff78-183">A variável `dialog` é declarada em um escopo mais amplo do que o retorno de chamada porque ela também é referenciada em `processMessage`.</span><span class="sxs-lookup"><span data-stu-id="3ff78-183">The `dialog` variable is declared at a wider scope than the callback because it is also referenced in `processMessage`.</span></span>

<span data-ttu-id="3ff78-184">Veja a seguir um exemplo simples de um manipulador para o evento `DialogMessageReceived`:</span><span class="sxs-lookup"><span data-stu-id="3ff78-184">The following is a simple example of a handler for the `DialogMessageReceived` event:</span></span>

```js
function processMessage(arg) {
    var messageFromDialog = JSON.parse(arg.message);
    showUserName(messageFromDialog.name);
}
```

> [!NOTE]
> - <span data-ttu-id="3ff78-185">O Office transmite o objeto `arg` para o manipulador.</span><span class="sxs-lookup"><span data-stu-id="3ff78-185">Office passes the `arg` object to the handler.</span></span> <span data-ttu-id="3ff78-186">Sua propriedade `message` é o booliano ou a cadeia de caracteres enviada pela chamada de `messageParent` na caixa de diálogo.</span><span class="sxs-lookup"><span data-stu-id="3ff78-186">Its `message` property is the Boolean or string sent by the call of `messageParent` in the dialog box.</span></span> <span data-ttu-id="3ff78-187">Neste exemplo, é uma representação em formato de um perfil de usuário de um serviço como a conta da Microsoft ou o Google, para que seja desserializado de volta para um objeto com `JSON.parse` .</span><span class="sxs-lookup"><span data-stu-id="3ff78-187">In this example, it is a stringified representation of a user's profile from a service such as Microsoft account or Google, so it is deserialized back to an object with `JSON.parse`.</span></span>
> - <span data-ttu-id="3ff78-p124">A implementação de `showUserName` não é mostrada. Ela pode exibir uma mensagem de boas-vindas personalizada no painel de tarefas.</span><span class="sxs-lookup"><span data-stu-id="3ff78-p124">The `showUserName` implementation is not shown. It might display a personalized welcome message on the task pane.</span></span>

<span data-ttu-id="3ff78-190">Quando a interação do usuário com a caixa de diálogo for concluída, seu manipulador de mensagem fechará a caixa de diálogo, conforme mostrado neste exemplo.</span><span class="sxs-lookup"><span data-stu-id="3ff78-190">When the user interaction with the dialog box is completed, your message handler should close the dialog box, as shown in this example.</span></span>

```js
function processMessage(arg) {
    dialog.close();
    // message processing code goes here;
}
```

> [!NOTE]
> - <span data-ttu-id="3ff78-191">O objeto `dialog` deve ser o mesmo que é retornado pela chamada de `displayDialogAsync`.</span><span class="sxs-lookup"><span data-stu-id="3ff78-191">The `dialog` object must be the same one that is returned by the call of `displayDialogAsync`.</span></span>
> - <span data-ttu-id="3ff78-192">A chamada de `dialog.close` informa ao Office para fechar a caixa de diálogo imediatamente.</span><span class="sxs-lookup"><span data-stu-id="3ff78-192">The call of `dialog.close` tells Office to immediately close the dialog box.</span></span>

<span data-ttu-id="3ff78-193">Para ver um suplemento de exemplo que usa essas técnicas, confira [Exemplo de API de Caixa de diálogo do Suplemento do Office](https://github.com/OfficeDev/Office-Add-in-Dialog-API-Simple-Example).</span><span class="sxs-lookup"><span data-stu-id="3ff78-193">For a sample add-in that uses these techniques, see [Office Add-in Dialog API Example](https://github.com/OfficeDev/Office-Add-in-Dialog-API-Simple-Example).</span></span>

<span data-ttu-id="3ff78-p125">Se o suplemento precisa abrir uma página diferente do painel de tarefas depois de receber a mensagem, é possível usar o método `window.location.replace` (ou `window.location.href`) como a última linha do manipulador. Apresentamos um exemplo a seguir:</span><span class="sxs-lookup"><span data-stu-id="3ff78-p125">If the add-in needs to open a different page of the task pane after receiving the message, you can use the `window.location.replace` method (or `window.location.href`) as the last line of the handler. The following is an example:</span></span>

```js
function processMessage(arg) {
    // message processing code goes here;
    window.location.replace("/newPage.html");
    // Alternatively ...
    // window.location.href = "/newPage.html";
}
```

<span data-ttu-id="3ff78-196">Para ver um exemplo de um suplemento que faz isso, consulte [Inserir gráficos do Excel usando o Microsoft Graph em um Suplemento do PowerPoint](https://github.com/OfficeDev/PowerPoint-Add-in-Microsoft-Graph-ASPNET-InsertChart).</span><span class="sxs-lookup"><span data-stu-id="3ff78-196">For an example of an add-in that does this, see the [Insert Excel charts using Microsoft Graph in a PowerPoint add-in](https://github.com/OfficeDev/PowerPoint-Add-in-Microsoft-Graph-ASPNET-InsertChart) sample.</span></span>

### <a name="conditional-messaging"></a><span data-ttu-id="3ff78-197">Mensagens condicionais</span><span class="sxs-lookup"><span data-stu-id="3ff78-197">Conditional messaging</span></span>

<span data-ttu-id="3ff78-198">Como você pode enviar várias chamadas `messageParent` a partir da caixa de diálogo, mas tem apenas um manipulador na página host do evento `DialogMessageReceived`, o manipulador tem que usar a lógica condicional para distinguir mensagens diferentes.</span><span class="sxs-lookup"><span data-stu-id="3ff78-198">Because you can send multiple `messageParent` calls from the dialog box, but you have only one handler in the host page for the `DialogMessageReceived` event, the handler must use conditional logic to distinguish different messages.</span></span> <span data-ttu-id="3ff78-199">Por exemplo, se a caixa de diálogo solicitar que um usuário entre em um provedor de identidade como a conta da Microsoft ou Google, ele enviará o perfil do usuário como uma mensagem.</span><span class="sxs-lookup"><span data-stu-id="3ff78-199">For example, if the dialog box prompts a user to sign in to an identity provider such as Microsoft account or Google, it sends the user's profile as a message.</span></span> <span data-ttu-id="3ff78-200">Se a autenticação falhar, a caixa de diálogo enviará informações de erro à página host, como no exemplo a seguir:</span><span class="sxs-lookup"><span data-stu-id="3ff78-200">If authentication fails, the dialog box sends error information to the host page, as in the following example:</span></span>

```js
if (loginSuccess) {
    var userProfile = getProfile();
    var messageObject = {messageType: "signinSuccess", profile: userProfile};
    var jsonMessage = JSON.stringify(messageObject);
    Office.context.ui.messageParent(jsonMessage);
} else {
    var errorDetails = getError();
    var messageObject = {messageType: "signinFailure", error: errorDetails};
    var jsonMessage = JSON.stringify(messageObject);
    Office.context.ui.messageParent(jsonMessage);
}
```

> [!NOTE]
> - <span data-ttu-id="3ff78-201">A variável `loginSuccess` poderia ser inicializada por meio da leitura da resposta HTTP no provedor de identidade.</span><span class="sxs-lookup"><span data-stu-id="3ff78-201">The `loginSuccess` variable would be initialized by reading the HTTP response from the identity provider.</span></span>
> - <span data-ttu-id="3ff78-p127">A implementação das funções `getProfile` e `getError` não é exibida. Cada uma delas obtém dados de um parâmetro de consulta ou do corpo da resposta HTTP.</span><span class="sxs-lookup"><span data-stu-id="3ff78-p127">The the implementation of the `getProfile` and `getError` functions are not not shown. They each get data from a query parameter or from the body of the HTTP response.</span></span>
> - <span data-ttu-id="3ff78-p128">São enviados objetos anônimos de diferentes tipos se a entrada for bem-sucedida ou não. Ambos têm uma propriedade `messageType`, mas um tem uma propriedade `profile` e o outro tem uma propriedade `error`.</span><span class="sxs-lookup"><span data-stu-id="3ff78-p128">Anonymous objects of different types are sent depending on whether the sign in was successful. Both have a `messageType` property, but one has a `profile` property and the other has an `error` property.</span></span>

<span data-ttu-id="3ff78-p129">O código do manipulador na página host usa o valor da propriedade `messageType` para ramificar como no exemplo a seguir. A função `showUserName` é a mesma do exemplo anterior e a função `showNotification` exibe o erro na interface do usuário da página host.</span><span class="sxs-lookup"><span data-stu-id="3ff78-p129">The handler code in the host page uses the value of the `messageType` property to branch as shown in the following example. Note that the `showUserName` function is the same as in the previous example and `showNotification` function displays the error in the host page's UI.</span></span>

```js
function processMessage(arg) {
    var messageFromDialog = JSON.parse(arg.message);
    if (messageFromDialog.messageType === "signinSuccess") {
        dialog.close();
        showUserName(messageFromDialog.profile.name);
        window.location.replace("/newPage.html");
    } else {
        dialog.close();
        showNotification("Unable to authenticate user: " + messageFromDialog.error);
    }
}
```

> [!NOTE]
> <span data-ttu-id="3ff78-208">A `showNotification` implementação não é exibida no código de exemplo fornecido neste artigo.</span><span class="sxs-lookup"><span data-stu-id="3ff78-208">The `showNotification` implementation is not shown in the sample code provided by this article.</span></span> <span data-ttu-id="3ff78-209">Um exemplo de como você pode implementar essa função dentro do suplemento, confira [Exemplo do suplemento do Office exemplo do diálogo API](https://github.com/OfficeDev/Office-Add-in-Dialog-API-Simple-Example).</span><span class="sxs-lookup"><span data-stu-id="3ff78-209">For an example of how you might implement this function within your add-in, see [Office Add-in Dialog API Example](https://github.com/OfficeDev/Office-Add-in-Dialog-API-Simple-Example).</span></span>

## <a name="pass-information-to-the-dialog-box"></a><span data-ttu-id="3ff78-210">Transmitir informações para a caixa diálogo</span><span class="sxs-lookup"><span data-stu-id="3ff78-210">Pass information to the dialog box</span></span>

<span data-ttu-id="3ff78-p131">Às vezes, a página host precisa transmitir informações para a caixa de diálogo. Você pode fazer isso de duas maneiras principais:</span><span class="sxs-lookup"><span data-stu-id="3ff78-p131">Sometimes the host page needs to pass information to the dialog box. You can do this in two primary ways:</span></span>

- <span data-ttu-id="3ff78-213">Adicionar parâmetros de consulta à URL que é transmitida para `displayDialogAsync`.</span><span class="sxs-lookup"><span data-stu-id="3ff78-213">Add query parameters to the URL that is passed to `displayDialogAsync`.</span></span>
- <span data-ttu-id="3ff78-214">Armazenar as informações em outro local que seja acessível para a janela do host e para a caixa de diálogo.</span><span class="sxs-lookup"><span data-stu-id="3ff78-214">Store the information somewhere that is accessible to both the host window and dialog box.</span></span> <span data-ttu-id="3ff78-215">As duas janelas não compartilham um armazenamento de sessão comum, mas *se elas tiverem o mesmo domínio* (incluindo o número da porta, se houver algum), compartilharão um [local de armazenamento](https://www.w3schools.com/html/html5_webstorage.asp) comum.\*</span><span class="sxs-lookup"><span data-stu-id="3ff78-215">The two windows do not share a common session storage, but *if they have the same domain* (including port number, if any), they share a common [Local Storage](https://www.w3schools.com/html/html5_webstorage.asp).\*</span></span>

> [!NOTE]
> <span data-ttu-id="3ff78-216">\* Há um bug que afetará sua estratégia de tratamento de tokens.</span><span class="sxs-lookup"><span data-stu-id="3ff78-216">\* There is a bug that will effect your strategy for token handling.</span></span> <span data-ttu-id="3ff78-217">Se o suplemento estiver sendo executado no **Office na Web** nos navegadores Safari ou Microsoft Edge, o painel de tarefas e a caixa de diálogo não compartilharão o mesmo Armazenamento Local, portanto, ele não poderá ser usado para a comunicação entre eles.</span><span class="sxs-lookup"><span data-stu-id="3ff78-217">If the add-in is running in **Office on the web** in either the Safari or Edge browser, the dialog box and task pane do not share the same Local Storage, so it cannot be used to communicate between them.</span></span>

### <a name="use-local-storage"></a><span data-ttu-id="3ff78-218">Usar o armazenamento local</span><span class="sxs-lookup"><span data-stu-id="3ff78-218">Use local storage</span></span>

<span data-ttu-id="3ff78-219">Para usar o armazenamento local, seu código chama o método `setItem` do objeto `window.localStorage` na página host antes da chamada `displayDialogAsync`, como no exemplo a seguir:</span><span class="sxs-lookup"><span data-stu-id="3ff78-219">To use local storage, your code calls the `setItem` method of the `window.localStorage` object in the host page before the `displayDialogAsync` call, as in the following example:</span></span>

```js
localStorage.setItem("clientID", "15963ac5-314f-4d9b-b5a1-ccb2f1aea248");
```

<span data-ttu-id="3ff78-220">O código na caixa de diálogo lê o item quando necessário, como no exemplo a seguir:</span><span class="sxs-lookup"><span data-stu-id="3ff78-220">Code in the dialog box reads the item when it's needed, as in the following example:</span></span>

```js
var clientID = localStorage.getItem("clientID");
// You can also use property syntax:
// var clientID = localStorage.clientID;
```

### <a name="use-query-parameters"></a><span data-ttu-id="3ff78-221">Usar parâmetros de consulta</span><span class="sxs-lookup"><span data-stu-id="3ff78-221">Use query parameters</span></span>

<span data-ttu-id="3ff78-222">O exemplo a seguir mostra como transmitir dados com um parâmetro de consulta:</span><span class="sxs-lookup"><span data-stu-id="3ff78-222">The following example shows how to pass data with a query parameter:</span></span>

```js
Office.context.ui.displayDialogAsync('https://myAddinDomain/myDialog.html?clientID=15963ac5-314f-4d9b-b5a1-ccb2f1aea248');
```

<span data-ttu-id="3ff78-223">Para ver um exemplo que usa essa técnica, consulte [Inserir gráficos do Excel usando o Microsoft Graph em um Suplemento do PowerPoint](https://github.com/OfficeDev/PowerPoint-Add-in-Microsoft-Graph-ASPNET-InsertChart).</span><span class="sxs-lookup"><span data-stu-id="3ff78-223">For a sample that uses this technique, see [Insert Excel charts using Microsoft Graph in a PowerPoint add-in](https://github.com/OfficeDev/PowerPoint-Add-in-Microsoft-Graph-ASPNET-InsertChart).</span></span>

<span data-ttu-id="3ff78-224">O código na caixa de diálogo pode analisar a URL e ler o valor do parâmetro.</span><span class="sxs-lookup"><span data-stu-id="3ff78-224">Code in your dialog box can parse the URL and read the parameter value.</span></span>

> [!IMPORTANT]
> <span data-ttu-id="3ff78-p134">O Office adiciona automaticamente um parâmetro de consulta chamado `_host_info` à URL que é transmitida para `displayDialogAsync`. Ele é anexado após os parâmetros de consulta personalizados, se houver algum. Ele não é anexado às URLs subsequentes para as quais a caixa de diálogo navega. No futuro, a Microsoft poderá alterar o conteúdo desse valor ou removê-lo completamente para que seu código não consiga lê-lo. O mesmo valor é adicionado ao armazenamento de sessão da caixa de diálogo. Novamente, *seu código não deve ler nem gravar esse valor*.</span><span class="sxs-lookup"><span data-stu-id="3ff78-p134">Office automatically adds a query parameter called `_host_info` to the URL that is passed to `displayDialogAsync`. (It is appended after your custom query parameters, if any. It is not appended to any subsequent URLs that the dialog box navigates to.) Microsoft may change the content of this value, or remove it entirely, in the future, so your code should not read it. The same value is added to the dialog box's session storage. Again, *your code should neither read nor write to this value*.</span></span>

> [!NOTE]
> <span data-ttu-id="3ff78-230">Agora, há uma `messageChild` API que a página pai pode usar para enviar mensagens para a caixa de diálogo, assim como a `messageParent` API descrita acima envia mensagens da caixa de diálogo.</span><span class="sxs-lookup"><span data-stu-id="3ff78-230">There is now in preview a `messageChild` API that the parent page can use to send messages to the dialog just as the `messageParent` API described above sends messages from the dialog.</span></span> <span data-ttu-id="3ff78-231">Para saber mais sobre ele, consulte [passando dados e mensagens para uma caixa de diálogo da página host](parent-to-dialog.md).</span><span class="sxs-lookup"><span data-stu-id="3ff78-231">For more about it, see [Passing data and messages to a dialog box from its host page](parent-to-dialog.md).</span></span> <span data-ttu-id="3ff78-232">Recomendamos que você experimente, mas para suplementos de produção, é recomendável usar as técnicas descritas nesta seção.</span><span class="sxs-lookup"><span data-stu-id="3ff78-232">We encourage you to try it out, but for production add-ins, we recommend that you use the techniques described in this section.</span></span>

## <a name="closing-the-dialog-box"></a><span data-ttu-id="3ff78-233">Feche a caixa de diálogo</span><span class="sxs-lookup"><span data-stu-id="3ff78-233">Closing the dialog box</span></span>

<span data-ttu-id="3ff78-p136">Você pode implementar um botão na caixa de diálogo para fechá-la. Para fazer isso, o manipulador de eventos de clique do botão deve usar `messageParent` para informar a página host em que o botão foi clicado. Apresentamos um exemplo a seguir:</span><span class="sxs-lookup"><span data-stu-id="3ff78-p136">You can implement a button in the dialog box that will close it. To do this, the click event handler for the button should use `messageParent` to tell the host page that the button has been clicked. The following is an example:</span></span>

```js
function closeButtonClick() {
    var messageObject = {messageType: "dialogClosed"};
    var jsonMessage = JSON.stringify(messageObject);
    Office.context.ui.messageParent(jsonMessage);
}
```

<span data-ttu-id="3ff78-237">O manipulador de página host de `DialogMessageReceived` poderia chamar `dialog.close`, como neste exemplo.</span><span class="sxs-lookup"><span data-stu-id="3ff78-237">The host page handler for `DialogMessageReceived` would call `dialog.close`, as in this example.</span></span> <span data-ttu-id="3ff78-238">(Veja exemplos anteriores que mostram como o objeto `dialog` é inicializado.)</span><span class="sxs-lookup"><span data-stu-id="3ff78-238">(See previous examples that show how the `dialog` object is initialized.)</span></span>

```js
function processMessage(arg) {
    var messageFromDialog = JSON.parse(arg.message);
    if (messageFromDialog.messageType === "dialogClosed") {
       dialog.close();
    }
}
```

<span data-ttu-id="3ff78-239">Mesmo quando você não tem sua própria interface de usuário de diálogo de fechar, um usuário final pode fechar a caixa de diálogo escolhendo a opção **X** no canto superior direito.</span><span class="sxs-lookup"><span data-stu-id="3ff78-239">Even when you don't have your own close-dialog UI, an end user can close the dialog box by choosing the **X** in the upper-right corner.</span></span> <span data-ttu-id="3ff78-240">Essa ação aciona o evento `DialogEventReceived`.</span><span class="sxs-lookup"><span data-stu-id="3ff78-240">This action triggers the `DialogEventReceived` event.</span></span> <span data-ttu-id="3ff78-241">Se seu painel do host precisar saber quando isso acontece, ele deverá declarar um manipulador para esse evento.</span><span class="sxs-lookup"><span data-stu-id="3ff78-241">If your host pane needs to know when this happens, it should declare a handler for this event.</span></span> <span data-ttu-id="3ff78-242">Confira a seção [Erros e eventos na caixa de diálogo](dialog-handle-errors-events.md#errors-and-events-in-the-dialog-box) para ver os detalhes.</span><span class="sxs-lookup"><span data-stu-id="3ff78-242">See the section [Errors and events in the dialog box](dialog-handle-errors-events.md#errors-and-events-in-the-dialog-box) for details.</span></span>

## <a name="advanced-topics-and-special-scenarios"></a><span data-ttu-id="3ff78-243">Tópicos avançados e cenários especiais</span><span class="sxs-lookup"><span data-stu-id="3ff78-243">Advanced topics and special scenarios</span></span>

### <a name="use-the-dialog-api-to-show-a-video"></a><span data-ttu-id="3ff78-244">Use a API de Caixa de Diálogo para exibir um vídeo</span><span class="sxs-lookup"><span data-stu-id="3ff78-244">Use the Dialog API to show a video</span></span>

<span data-ttu-id="3ff78-245">Confira [use a caixa de diálogo do Office para mostrar um vídeo](dialog-video.md).</span><span class="sxs-lookup"><span data-stu-id="3ff78-245">See [Use the Office dialog box to show a video](dialog-video.md).</span></span>

### <a name="use-the-dialog-apis-in-an-authentication-flow"></a><span data-ttu-id="3ff78-246">Use as APIs de Caixa de Diálogo em um fluxo de autenticação</span><span class="sxs-lookup"><span data-stu-id="3ff78-246">Use the Dialog APIs in an authentication flow</span></span>

<span data-ttu-id="3ff78-247">Confira[Autenticar com a API da Caixa de Diálogo do Office](auth-with-office-dialog-api.md).</span><span class="sxs-lookup"><span data-stu-id="3ff78-247">See [Authenticate with the Office dialog API](auth-with-office-dialog-api.md).</span></span>

### <a name="using-the-office-dialog-api-with-single-page-applications-and-client-side-routing"></a><span data-ttu-id="3ff78-248">Usar a API de Caixa de diálogo do Office com aplicativos de página única e roteamento do lado do cliente</span><span class="sxs-lookup"><span data-stu-id="3ff78-248">Using the Office dialog API with single-page applications and client-side routing</span></span>

<span data-ttu-id="3ff78-249">SPAs e o roteamento do lado do cliente devem ser manuseados com cuidado ao usar a API de diálogo do Office.</span><span class="sxs-lookup"><span data-stu-id="3ff78-249">SPAs and client-side routing need to be handled with care when you are using the Office dialog API.</span></span> <span data-ttu-id="3ff78-250">Confira [práticas recomendadas para usar o Office Dialog API em um SPA](dialog-best-practices.md#best-practices-for-using-the-office-dialog-api-in-an-spa).</span><span class="sxs-lookup"><span data-stu-id="3ff78-250">Please see [Best practices for using the Office dialog API in an SPA](dialog-best-practices.md#best-practices-for-using-the-office-dialog-api-in-an-spa).</span></span>

### <a name="error-and-event-handling"></a><span data-ttu-id="3ff78-251">Manipulação de erros e eventos</span><span class="sxs-lookup"><span data-stu-id="3ff78-251">Error and event handling</span></span>

<span data-ttu-id="3ff78-252">Confira [Manipulando erros e eventos na caixa de diálogo do Office](dialog-handle-errors-events.md).</span><span class="sxs-lookup"><span data-stu-id="3ff78-252">See [Handling errors and events in the Office dialog box](dialog-handle-errors-events.md).</span></span>

## <a name="next-steps"></a><span data-ttu-id="3ff78-253">Próximas etapas</span><span class="sxs-lookup"><span data-stu-id="3ff78-253">Next steps</span></span>

<span data-ttu-id="3ff78-254">Saiba mais sobre as armadilhas e as práticas recomendadas para a API de diálogo do Office em [práticas recomendadas e regras para a API do Office Dialog](dialog-best-practices.md).</span><span class="sxs-lookup"><span data-stu-id="3ff78-254">Learn about gotchas and best practices for the Office dialog API in [Best practices and rules for the Office dialog API](dialog-best-practices.md).</span></span>
