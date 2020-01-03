---
title: Use a API de Caixa de Diálogo em seus Suplementos do Office
description: ''
ms.date: 08/07/2019
localization_priority: Priority
ms.openlocfilehash: 88c7afca2f1e800391443458e0c6f6b930288c44
ms.sourcegitcommit: 8c5c5a1bd3fe8b90f6253d9850e9352ed0b283ee
ms.translationtype: HT
ms.contentlocale: pt-BR
ms.lasthandoff: 12/19/2019
ms.locfileid: "40814107"
---
# <a name="use-the-dialog-api-in-your-office-add-ins"></a><span data-ttu-id="cdb98-102">Use a API de Caixa de Diálogo em seus Suplementos do Office</span><span class="sxs-lookup"><span data-stu-id="cdb98-102">Use the Dialog API in your Office Add-ins</span></span>

<span data-ttu-id="cdb98-p101">Você pode usar a [API de Caixa de diálogo](/javascript/api/office/office.ui) para abrir caixas de diálogo no seu Suplemento do Office. Este artigo fornece orientações para usar a API de Caixa de diálogo em seu Suplemento do Office.</span><span class="sxs-lookup"><span data-stu-id="cdb98-p101">You can use the [Dialog API](/javascript/api/office/office.ui) to open dialog boxes in your Office Add-in. This article provides guidance for using the Dialog API in your Office Add-in.</span></span>

> [!NOTE]
> <span data-ttu-id="cdb98-p102">Para informações sobre os programas para os quais a API de Caixa de Diálogo tem suporte no momento, confira [Conjuntos de requisitos da API de Caixa de Diálogo](/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets). Atualmente, a API de Caixa de Diálogo tem suporte para Word, Excel, PowerPoint e Outlook.</span><span class="sxs-lookup"><span data-stu-id="cdb98-p102">For information about where the Dialog API is currently supported, see [Dialog API requirement sets](/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets). The Dialog API is currently supported for Word, Excel, PowerPoint, and Outlook.</span></span>

<span data-ttu-id="cdb98-107">Um cenário fundamental para as APIs de Caixa de Diálogo é habilitar a autenticação com um recurso como o Google, o Facebook ou o Microsoft Graph.</span><span class="sxs-lookup"><span data-stu-id="cdb98-107">A primary scenario for the Dialog APIs is to enable authentication with a resource such as Google, Facebook, or Microsoft Graph.</span></span> <span data-ttu-id="cdb98-108">Para saber mais, confira [ autenticação com APIs de Caixa de Diálogo do Office](auth-with-office-dialog-api.md) *depois* que você se familiarizar com este artigo.</span><span class="sxs-lookup"><span data-stu-id="cdb98-108">For more information, see [Authenticate with the Office Dialog API](auth-with-office-dialog-api.md) *after* you are familiar with this article.</span></span>

<span data-ttu-id="cdb98-109">Considere abrir uma caixa de diálogo em um painel de tarefas, suplemento de conteúdo ou [comando de suplemento](../design/add-in-commands.md) para fazer o seguinte:</span><span class="sxs-lookup"><span data-stu-id="cdb98-109">Consider opening a dialog box from a task pane or content add-in or [add-in command](../design/add-in-commands.md) to do the following:</span></span>

- <span data-ttu-id="cdb98-110">Exibir páginas de entrada que não podem ser abertas diretamente em um painel de tarefas.</span><span class="sxs-lookup"><span data-stu-id="cdb98-110">Display sign in pages that cannot be opened directly in a task pane.</span></span>
- <span data-ttu-id="cdb98-111">Fornecer mais espaço na tela, ou até uma tela inteira, para algumas tarefas no seu suplemento.</span><span class="sxs-lookup"><span data-stu-id="cdb98-111">Provide more screen space, or even a full screen, for some tasks in your add-in.</span></span>
- <span data-ttu-id="cdb98-112">Hospedar um vídeo que seria muito pequeno se fosse confinado em um painel de tarefas.</span><span class="sxs-lookup"><span data-stu-id="cdb98-112">Host a video that would be too small if confined to a task pane.</span></span>

> [!NOTE]
> <span data-ttu-id="cdb98-p104">Como a sobreposição de elementos de IU não são recomendáveis, evite abrir uma caixa de diálogo em um painel de tarefas a menos que seu cenário o obrigue a fazer isso. Ao considerar como usar a área de superfície de um painel de tarefas, observe que painéis de tarefas podem ter guias. Por exemplo, confira o exemplo [Suplemento do Excel JavaScriptSalesTracker](https://github.com/OfficeDev/Excel-Add-in-JavaScript-SalesTracker).</span><span class="sxs-lookup"><span data-stu-id="cdb98-p104">Because overlapping UI elements are discouraged, avoid opening a dialog from a task pane unless your scenario requires it. When you consider how to use the surface area of a task pane, note that task panes can be tabbed. For an example, see the [Excel Add-in JavaScript SalesTracker](https://github.com/OfficeDev/Excel-Add-in-JavaScript-SalesTracker) sample.</span></span>

<span data-ttu-id="cdb98-116">A imagem abaixo mostra um exemplo de uma caixa de diálogo.</span><span class="sxs-lookup"><span data-stu-id="cdb98-116">The following image shows an example of a dialog box.</span></span>

![Comandos de suplemento](../images/auth-o-dialog-open.png)

<span data-ttu-id="cdb98-p105">A caixa de diálogo sempre abre no centro da tela. O usuário pode movê-la e redimensioná-la. A janela é *não modal*: o usuário pode continuar a interagir com o documento no aplicativo do Office do host e com a página host no painel de tarefas, caso houver uma.</span><span class="sxs-lookup"><span data-stu-id="cdb98-p105">Note that the dialog box always opens in the center of the screen. The user can move and resize it. The window is *nonmodal*--a user can continue to interact with both the document in the host Office application and with the host page in the task pane, if there is one.</span></span>

## <a name="dialog-api-scenarios"></a><span data-ttu-id="cdb98-121">Cenários da API de Caixa de Diálogo</span><span class="sxs-lookup"><span data-stu-id="cdb98-121">Dialog API scenarios</span></span>

<span data-ttu-id="cdb98-122">As APIs JavaScript para Office têm suporte para os seguintes cenários com um objeto [Dialog](/javascript/api/office/office.dialog) e duas funções no [namespace Office.context.ui](/javascript/api/office/office.ui).</span><span class="sxs-lookup"><span data-stu-id="cdb98-122">The Office JavaScript APIs support the following scenarios with a [Dialog](/javascript/api/office/office.dialog) object and two functions in the [Office.context.ui namespace](/javascript/api/office/office.ui).</span></span>

### <a name="open-a-dialog-box"></a><span data-ttu-id="cdb98-123">Abra uma caixa de diálogo</span><span class="sxs-lookup"><span data-stu-id="cdb98-123">Open a dialog box</span></span>

<span data-ttu-id="cdb98-p106">Para abrir uma caixa de diálogo, seu código no painel de tarefas chama o método [displayDialogAsync](/javascript/api/office/office.ui) e transmite a ele a URL do recurso que você deseja abrir. Isso geralmente é uma página, mas pode ser um método controlador em um aplicativo MVC, uma rota, um método de serviço Web ou qualquer outro recurso. Neste artigo, 'página' ou 'site' refere-se ao recurso na caixa de diálogo. Apresentamos um exemplo de código simples a seguir.</span><span class="sxs-lookup"><span data-stu-id="cdb98-p106">To open a dialog box, your code in the task pane calls the [displayDialogAsync](/javascript/api/office/office.ui) method and passes to it the URL of the resource that you want to open. This is usually a page, but it can be a controller method in an MVC application, a route, a web service method, or any other resource. In this article, 'page' or 'website' refers to the resource in the dialog. The following code is a simple example:</span></span>

```js
Office.context.ui.displayDialogAsync('https://myAddinDomain/myDialog.html');
```

> [!NOTE]
> - <span data-ttu-id="cdb98-p107">A URL usa o protocolo HTTP**S**. Isso é obrigatório para todas as páginas carregadas em uma caixa diálogo, não apenas para a primeira página carregada.</span><span class="sxs-lookup"><span data-stu-id="cdb98-p107">The URL uses the HTTP**S** protocol. This is mandatory for all pages loaded in a dialog box, not just the first page loaded.</span></span>
> - <span data-ttu-id="cdb98-130">O domínio do recurso de caixa de diálogo é o mesmo que o domínio da página de host, que pode ser a página em um painel de tarefas ou o [arquivo de função](/office/dev/add-ins/reference/manifest/functionfile) de um comando de suplemento.</span><span class="sxs-lookup"><span data-stu-id="cdb98-130">The dialog resource's domain is the same as the domain of the host page, which can be the page in a task pane or the [function file](/office/dev/add-ins/reference/manifest/functionfile) of an add-in command.</span></span> <span data-ttu-id="cdb98-131">Isso é necessário: a página, o método do controlador ou outro recurso que é passado para o método `displayDialogAsync` deve estar no mesmo domínio que a página de host.</span><span class="sxs-lookup"><span data-stu-id="cdb98-131">This is required: the page, controller method, or other resource that is passed to the `displayDialogAsync` method must be in the same domain as the host page.</span></span>

> [!IMPORTANT]
> <span data-ttu-id="cdb98-132">A página de host e os recursos de caixa de diálogo devem ter o mesmo domínio completo.</span><span class="sxs-lookup"><span data-stu-id="cdb98-132">The host page and the resources of the dialog must have the same full domain.</span></span> <span data-ttu-id="cdb98-133">Se você tentar passar `displayDialogAsync` para um subdomínio do domínio do suplemento, ele não funcionará.</span><span class="sxs-lookup"><span data-stu-id="cdb98-133">If you attempt to pass `displayDialogAsync` a subdomain of the add-in's domain, it will not work.</span></span> <span data-ttu-id="cdb98-134">O domínio completo, incluindo qualquer subdomínio, deve corresponder.</span><span class="sxs-lookup"><span data-stu-id="cdb98-134">The full domain, including any subdomain, must match.</span></span>

<span data-ttu-id="cdb98-p110">Após o carregamento da primeira página (ou de outro recurso), um usuário pode ir para qualquer site (ou outro recurso) que usa HTTPS. Também é possível criar a primeira página para redirecionar imediatamente para outro site.</span><span class="sxs-lookup"><span data-stu-id="cdb98-p110">After the first page (or other resource) is loaded, a user can go to any website (or other resource) that uses HTTPS. You can also design the first page to immediately redirect to another site.</span></span>

<span data-ttu-id="cdb98-137">Por padrão, a caixa de diálogo ocupará 80% da altura e da largura na tela do dispositivo, mas você pode definir porcentagens diferentes. Basta transmitir um objeto de configuração para o método, como mostra o exemplo a seguir.</span><span class="sxs-lookup"><span data-stu-id="cdb98-137">By default, the dialog box will occupy 80% of the height and width of the device screen, but you can set different percentages by passing a configuration object to the method, as shown in the following example:</span></span>

```js
Office.context.ui.displayDialogAsync('https://myDomain/myDialog.html', {height: 30, width: 20});
```

<span data-ttu-id="cdb98-138">Para ver um suplemento de exemplo que faz isso, confira [Exemplo de API de Caixa de diálogo do Suplemento do Office](https://github.com/OfficeDev/Office-Add-in-Dialog-API-Simple-Example).</span><span class="sxs-lookup"><span data-stu-id="cdb98-138">For a sample add-in that does this, see [Office Add-in Dialog API Example](https://github.com/OfficeDev/Office-Add-in-Dialog-API-Simple-Example).</span></span>

<span data-ttu-id="cdb98-p111">Defina os dois valores como 100% para ter uma verdadeira experiência de tela inteira. O máximo real é 99,5%, e a janela ainda poderá ser movida e redimensionada.</span><span class="sxs-lookup"><span data-stu-id="cdb98-p111">Set both values to 100% to get what is effectively a full screen experience. (The effective maximum is 99.5%, and the window is still moveable and resizable.)</span></span>

> [!NOTE]
> <span data-ttu-id="cdb98-p112">Apenas uma caixa de diálogo pode ser aberta em uma janela do host. Tentar abrir outra caixa de diálogo gera um erro. Portanto, por exemplo, se um usuário abrir uma caixa de diálogo no painel de tarefas, ele não poderá abrir uma segunda caixa de diálogo em uma página diferente no painel de tarefas. No entanto, quando uma caixa de diálogo é aberta em um [comando de suplemento](../design/add-in-commands.md), o comando abre um arquivo HTML novo (mas não visto) sempre que ele é selecionado. Isso cria uma nova janela do host (não vista) para que cada janela possa iniciar sua própria caixa de diálogo. Para obter mais informações, confira [Erros de displayDialogAsync](#errors-from-displaydialogasync).</span><span class="sxs-lookup"><span data-stu-id="cdb98-p112">You can open only one dialog box from a host window. An attempt to open another dialog box generates an error. For example, if a user opens a dialog box from a task pane, she cannot open a second dialog box, from a different page in the task pane. However, when a dialog box is opened from an [add-in command](../design/add-in-commands.md), the command opens a new (but unseen) HTML file each time it is selected. This creates a new (unseen) host window, so each such window can launch its own dialog box. For more information, see [Errors from displayDialogAsync](#errors-from-displaydialogasync).</span></span>

### <a name="take-advantage-of-a-performance-option-in-office-on-the-web"></a><span data-ttu-id="cdb98-147">Aproveite uma opção de desempenho no Office na Web</span><span class="sxs-lookup"><span data-stu-id="cdb98-147">Take advantage of a performance option in Office on the web</span></span>

<span data-ttu-id="cdb98-148">A propriedade `displayInIframe` é uma propriedade adicional no objeto de configuração que você pode passar para o`displayDialogAsync`.</span><span class="sxs-lookup"><span data-stu-id="cdb98-148">The `displayInIframe` property is an additional property in the configuration object that you can pass to `displayDialogAsync`.</span></span> <span data-ttu-id="cdb98-149">Quando essa propriedade for definida como `true` e o suplemento estiver em execução em um documento aberto no Office Online, a caixa de diálogo será aberta como um iframe flutuante, em vez de uma janela independente, o que faz com que ela seja aberta mais rapidamente.</span><span class="sxs-lookup"><span data-stu-id="cdb98-149">When this property is set to `true`, and the add-in is running in a document opened in Office on the web, the dialog box will open as a floating iframe rather than an independent window, which makes it open faster.</span></span> <span data-ttu-id="cdb98-150">Este é um exemplo:</span><span class="sxs-lookup"><span data-stu-id="cdb98-150">The following is an example:</span></span>

```js
Office.context.ui.displayDialogAsync('https://myDomain/myDialog.html', {height: 30, width: 20, displayInIframe: true});
```

<span data-ttu-id="cdb98-151">O valor padrão é `false`, que é o mesmo que omitir a propriedade inteiramente.</span><span class="sxs-lookup"><span data-stu-id="cdb98-151">The default value is `false`, which is the same as omitting the property entirely.</span></span> <span data-ttu-id="cdb98-152">Se o suplemento não estiver sendo executado no Office Online, o `displayInIframe` será ignorado.</span><span class="sxs-lookup"><span data-stu-id="cdb98-152">If the add-in is not running in Office on the web, the `displayInIframe` is ignored.</span></span>

> [!NOTE]
> <span data-ttu-id="cdb98-p115">Você **não** deverá usar `displayInIframe: true` se a caixa de diálogo redirecionar a qualquer ponto para uma página que não possa ser aberta em um iframe. Por exemplo, as páginas de entrada de muitos serviços Web populares, como Google e Conta da Microsoft, não podem ser abertas em um iframe.</span><span class="sxs-lookup"><span data-stu-id="cdb98-p115">You should **not** use `displayInIframe: true` if the dialog will at any point redirect to a page that cannot be opened in an iframe. For example, the sign in pages of many popular web services, such as Google and Microsoft Account, cannot be opened in an iframe.</span></span>

### <a name="handling-pop-up-blockers-with-office-on-the-web"></a><span data-ttu-id="cdb98-155">Tratamento de bloqueadores de pop-up com o Office na Web</span><span class="sxs-lookup"><span data-stu-id="cdb98-155">Handling pop-up blockers with Office on the web</span></span>

<span data-ttu-id="cdb98-156">Tentar exibir uma caixa de diálogo enquanto usa o Office na Web pode fazer com que bloqueadores de pop-up do navegador bloqueiem a caixa de diálogo.</span><span class="sxs-lookup"><span data-stu-id="cdb98-156">Attempting to display a dialog while using Office on the web may cause the browser's pop-up blocker to block the dialog.</span></span> <span data-ttu-id="cdb98-157">O bloqueador de pop-up do navegador pode ser evitado se o usuário de seu suplemento concordar primeiro com um aviso do suplemento.</span><span class="sxs-lookup"><span data-stu-id="cdb98-157">The browser's pop-up blocker can be circumvented if the user of your add-in first agrees to a prompt from the add-in.</span></span> <span data-ttu-id="cdb98-158">`displayDialogAsync`'s [DialogOptions](/javascript/api/office/office.dialogoptions) tem a `promptBeforeOpen` propriedade para acionar esse tipo de pop-up.</span><span class="sxs-lookup"><span data-stu-id="cdb98-158">`displayDialogAsync`'s [DialogOptions](/javascript/api/office/office.dialogoptions) has the `promptBeforeOpen` property to trigger such a pop-up.</span></span> <span data-ttu-id="cdb98-159">`promptBeforeOpen` é um valor booliano que fornece o comportamento a seguir:</span><span class="sxs-lookup"><span data-stu-id="cdb98-159">`promptBeforeOpen` is a boolean value which provides the following behavior:</span></span>

 - <span data-ttu-id="cdb98-160">`true` - A estrutura exibe um pop-up para acionar o painel de navegação e evitar bloqueadores de pop-up do navegador.</span><span class="sxs-lookup"><span data-stu-id="cdb98-160">`true` - The framework displays a pop-up to trigger the navigation and avoid the browser's pop-up blocker.</span></span> 
 - <span data-ttu-id="cdb98-161">`false` - A caixa de diálogo não será exibida e o desenvolvedor deverá lidar com pop-ups (fornecendo um artefato da interface de usuário para acionar a navegação).</span><span class="sxs-lookup"><span data-stu-id="cdb98-161">`false` - The dialog will not be shown and the developer must handle pop-ups (by providing a user interface artifact to trigger the navigation).</span></span> 
 
<span data-ttu-id="cdb98-162">O pop-up parece semelhante ao da captura de tela a seguir:</span><span class="sxs-lookup"><span data-stu-id="cdb98-162">The pop-up looks similiar to that in the following screenshot:</span></span>

![O aviso que uma caixa de diálogo do suplemento pode gerar para evitar bloqueadores de pop-up no navegador.](../images/dialog-prompt-before-open.png)
 
### <a name="send-information-from-the-dialog-box-to-the-host-page"></a><span data-ttu-id="cdb98-164">Envie informações da caixa de diálogo para a página host</span><span class="sxs-lookup"><span data-stu-id="cdb98-164">Send information from the dialog box to the host page</span></span>

<span data-ttu-id="cdb98-165">A caixa de diálogo não pode se comunicar com a página host no painel de tarefas, a menos que:</span><span class="sxs-lookup"><span data-stu-id="cdb98-165">The dialog box cannot communicate with the host page in the task pane unless:</span></span>

- <span data-ttu-id="cdb98-166">A página atual na caixa de diálogo esteja no mesmo domínio da página host.</span><span class="sxs-lookup"><span data-stu-id="cdb98-166">The current page in the dialog box is in the same domain as the host page.</span></span>
- <span data-ttu-id="cdb98-p117">A biblioteca JavaScript do Office seja carregada na página. Como qualquer página que usa a biblioteca JavaScript do Office, o script da página deve atribuir um método à propriedade `Office.initialize`, embora ele possa ser um método vazio. Para mais detalhes, confira [Iniciar o suplemento](understanding-the-javascript-api-for-office.md#initializing-your-add-in).</span><span class="sxs-lookup"><span data-stu-id="cdb98-p117">The Office JavaScript library is loaded in the page. (Like any page that uses the Office JavaScript library, script for the page must assign a method to the `Office.initialize` property, although it can be an empty method. For details, see [Initializing your add-in](understanding-the-javascript-api-for-office.md#initializing-your-add-in).)</span></span>

<span data-ttu-id="cdb98-p118">O código na página de diálogo use a função `messageParent` para enviar uma mensagem de cadeia de caracteres ou um valor booliano para a página host. A cadeia de caracteres pode ser uma palavra, uma frase, um blob XML, um JSON em formato de cadeia de caracteres ou qualquer outra coisa que possa ser serializada em uma cadeia de caracteres. Apresentamos um exemplo a seguir:</span><span class="sxs-lookup"><span data-stu-id="cdb98-p118">Code in the dialog page uses the `messageParent` function to send either a Boolean value or a string message to the host page. The string can be a word, sentence, XML blob, stringified JSON, or anything else that can be serialized to a string. The following is an example:</span></span>

```js
if (loginSuccess) {
    Office.context.ui.messageParent(true);
}
```

> [!NOTE]
> - <span data-ttu-id="cdb98-p119">A função `messageParent` é uma das *únicas* duas APIs do Office que pode ser chamada na caixa de diálogo. A outra é `Office.context.requirements.isSetSupported`. Para saber mais, confira [Especificar hosts do Office e requisitos da API](specify-office-hosts-and-api-requirements.md).</span><span class="sxs-lookup"><span data-stu-id="cdb98-p119">The `messageParent` function is one of *only* two Office APIs that can be called in the dialog box. The other is `Office.context.requirements.isSetSupported`. For information about it, see [Specify Office hosts and API requirements](specify-office-hosts-and-api-requirements.md).</span></span>
> - <span data-ttu-id="cdb98-176">A função `messageParent` só pode ser chamada em uma página com o mesmo domínio (incluindo o protocolo e a porta) da página host.</span><span class="sxs-lookup"><span data-stu-id="cdb98-176">The `messageParent` function can only be called on a page with the same domain (including protocol and port) as the host page.</span></span>

<span data-ttu-id="cdb98-177">No próximo exemplo, `googleProfile` é uma versão em formato de cadeia de caracteres do perfil do Google do usuário.</span><span class="sxs-lookup"><span data-stu-id="cdb98-177">In the next example, `googleProfile` is a stringified version of the user's Google profile.</span></span>

```js
if (loginSuccess) {
    Office.context.ui.messageParent(googleProfile);
}
```

<span data-ttu-id="cdb98-p120">A página host deve ser configurada para receber a mensagem. Você pode fazer isso adicionando um parâmetro de retorno de chamada à chamada original de `displayDialogAsync`. O retorno de chamada atribui um manipulador ao evento `DialogMessageReceived`. Apresentamos um exemplo a seguir:</span><span class="sxs-lookup"><span data-stu-id="cdb98-p120">The host page must be configured to receive the message. You do this by adding a callback parameter to the original call of `displayDialogAsync`. The callback assigns a handler to the `DialogMessageReceived` event. The following is an example:</span></span>

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
> - <span data-ttu-id="cdb98-p121">O Office transmite um objeto [AsyncResult](/javascript/api/office/office.asyncresult) para o retorno de chamada. Ele representa o resultado de tentativas de abrir a caixa de diálogo, mas não representa o resultado de eventos na caixa diálogo. Para obter mais informações sobre essa distinção, confira a seção [Manipular erros e eventos](#handle-errors-and-events).</span><span class="sxs-lookup"><span data-stu-id="cdb98-p121">Office passes an [AsyncResult](/javascript/api/office/office.asyncresult) object to the callback. It represents the result of the attempt to open the dialog box. It does not represent the outcome of any events in the dialog box. For more on this distinction, see the section [Handle errors and events](#handle-errors-and-events).</span></span>
> - <span data-ttu-id="cdb98-186">A propriedade `value` do `asyncResult` é definida como um objeto [Dialog](/javascript/api/office/office.dialog) que existe na página host, não no contexto da execução da caixa de diálogo.</span><span class="sxs-lookup"><span data-stu-id="cdb98-186">The `value` property of the `asyncResult` is set to a [Dialog](/javascript/api/office/office.dialog) object, which exists in the host page, not in the dialog box's execution context.</span></span>
> - <span data-ttu-id="cdb98-p122">O `processMessage` é a função que manipula o evento. Você pode dar a ele o nome que desejar.</span><span class="sxs-lookup"><span data-stu-id="cdb98-p122">The `processMessage` is the function that handles the event. You can give it any name you want.</span></span>
> - <span data-ttu-id="cdb98-189">A variável `dialog` é declarada em um escopo mais amplo do que o retorno de chamada porque ela também é referenciada em `processMessage`.</span><span class="sxs-lookup"><span data-stu-id="cdb98-189">The `dialog` variable is declared at a wider scope than the callback because it is also referenced in `processMessage`.</span></span>

<span data-ttu-id="cdb98-190">Veja a seguir um exemplo simples de um manipulador para o evento `DialogMessageReceived`:</span><span class="sxs-lookup"><span data-stu-id="cdb98-190">The following is a simple example of a handler for the `DialogMessageReceived` event:</span></span>

```js
function processMessage(arg) {
    var messageFromDialog = JSON.parse(arg.message);
    showUserName(messageFromDialog.name);
}
```

> [!NOTE]
> - <span data-ttu-id="cdb98-p123">O Office transmite o objeto `arg` para o manipulador. Sua propriedade `message` é o booliano ou a cadeia de caracteres enviada pela chamada de `messageParent` na caixa de diálogo. Neste exemplo, ela é uma representação em formato de cadeia de caracteres de um perfil de usuário de um serviço como a Conta da Microsoft ou o Google, portanto está desserializada como um objeto com `JSON.parse` novamente.</span><span class="sxs-lookup"><span data-stu-id="cdb98-p123">Office passes the `arg` object to the handler. Its `message` property is the Boolean or string sent by the call of `messageParent` in the dialog. In this example, it is a stringified representation of a user's profile from a service such as Microsoft Account or Google, so it is deserialized back to an object with `JSON.parse`.</span></span>
> - <span data-ttu-id="cdb98-p124">A implementação de `showUserName` não é mostrada. Ela pode exibir uma mensagem de boas-vindas personalizada no painel de tarefas.</span><span class="sxs-lookup"><span data-stu-id="cdb98-p124">The `showUserName` implementation is not shown. It might display a personalized welcome message on the task pane.</span></span>

<span data-ttu-id="cdb98-196">Quando a interação do usuário com a caixa de diálogo for concluída, seu manipulador de mensagem fechará a caixa de diálogo, conforme mostrado neste exemplo.</span><span class="sxs-lookup"><span data-stu-id="cdb98-196">When the user interaction with the dialog box is completed, your message handler should close the dialog box, as shown in this example.</span></span>

```js
function processMessage(arg) {
    dialog.close();
    // message processing code goes here;
}
```

> [!NOTE]
> - <span data-ttu-id="cdb98-197">O objeto `dialog` deve ser o mesmo que é retornado pela chamada de `displayDialogAsync`.</span><span class="sxs-lookup"><span data-stu-id="cdb98-197">The `dialog` object must be the same one that is returned by the call of `displayDialogAsync`.</span></span>
> - <span data-ttu-id="cdb98-198">A chamada de `dialog.close` informa ao Office para fechar a caixa de diálogo imediatamente.</span><span class="sxs-lookup"><span data-stu-id="cdb98-198">The call of `dialog.close` tells Office to immediately close the dialog box.</span></span>

<span data-ttu-id="cdb98-199">Para ver um suplemento de exemplo que usa essas técnicas, confira [Exemplo de API de Caixa de diálogo do Suplemento do Office](https://github.com/OfficeDev/Office-Add-in-Dialog-API-Simple-Example).</span><span class="sxs-lookup"><span data-stu-id="cdb98-199">For a sample add-in that uses these techniques, see [Office Add-in Dialog API Example](https://github.com/OfficeDev/Office-Add-in-Dialog-API-Simple-Example).</span></span>

<span data-ttu-id="cdb98-p125">Se o suplemento precisa abrir uma página diferente do painel de tarefas depois de receber a mensagem, é possível usar o método `window.location.replace` (ou `window.location.href`) como a última linha do manipulador. Apresentamos um exemplo a seguir:</span><span class="sxs-lookup"><span data-stu-id="cdb98-p125">If the add-in needs to open a different page of the task pane after receiving the message, you can use the `window.location.replace` method (or `window.location.href`) as the last line of the handler. The following is an example:</span></span>

```js
function processMessage(arg) {
    // message processing code goes here;
    window.location.replace("/newPage.html");
    // Alternatively ...
    // window.location.href = "/newPage.html";
}
```

<span data-ttu-id="cdb98-202">Para ver um exemplo de um suplemento que faz isso, consulte [Inserir gráficos do Excel usando o Microsoft Graph em um Suplemento do PowerPoint](https://github.com/OfficeDev/PowerPoint-Add-in-Microsoft-Graph-ASPNET-InsertChart).</span><span class="sxs-lookup"><span data-stu-id="cdb98-202">For an example of an add-in that does this, see the [Insert Excel charts using Microsoft Graph in a PowerPoint add-in](https://github.com/OfficeDev/PowerPoint-Add-in-Microsoft-Graph-ASPNET-InsertChart) sample.</span></span>

#### <a name="conditional-messaging"></a><span data-ttu-id="cdb98-203">Mensagens condicionais</span><span class="sxs-lookup"><span data-stu-id="cdb98-203">Conditional messaging</span></span>

<span data-ttu-id="cdb98-p126">Como você pode enviar várias chamadas `messageParent` a partir da caixa de diálogo, mas tem apenas um manipulador na página host do evento `DialogMessageReceived`, o manipulador tem que usar a lógica condicional para distinguir mensagens diferentes. Por exemplo, se a caixa de diálogo solicitar que o usuário entre em um provedor de identidade como a Conta da Microsoft ou o Google, ele enviará o perfil do usuário como uma mensagem. Se a autenticação falhar, a caixa de diálogo enviará informações de erro à página host, como no exemplo a seguir:</span><span class="sxs-lookup"><span data-stu-id="cdb98-p126">Because you can send multiple `messageParent` calls from the dialog box, but you have only one handler in the host page for the `DialogMessageReceived` event, the handler must use conditional logic to distinguish different messages. For example, if the dialog box prompts a user to sign in to an identity provider such as Microsoft Account or Google, it sends the user's profile as a message. If authentication fails, the dialog box sends error information to the host page, as in the following example:</span></span>

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
> - <span data-ttu-id="cdb98-207">A variável `loginSuccess` poderia ser inicializada por meio da leitura da resposta HTTP no provedor de identidade.</span><span class="sxs-lookup"><span data-stu-id="cdb98-207">The `loginSuccess` variable would be initialized by reading the HTTP response from the identity provider.</span></span>
> - <span data-ttu-id="cdb98-p127">A implementação das funções `getProfile` e `getError` não é exibida. Cada uma delas obtém dados de um parâmetro de consulta ou do corpo da resposta HTTP.</span><span class="sxs-lookup"><span data-stu-id="cdb98-p127">The the implementation of the `getProfile` and `getError` functions are not not shown. They each get data from a query parameter or from the body of the HTTP response.</span></span>
> - <span data-ttu-id="cdb98-p128">São enviados objetos anônimos de diferentes tipos se a entrada for bem-sucedida ou não. Ambos têm uma propriedade `messageType`, mas um tem uma propriedade `profile` e o outro tem uma propriedade `error`.</span><span class="sxs-lookup"><span data-stu-id="cdb98-p128">Anonymous objects of different types are sent depending on whether the sign in was successful. Both have a `messageType` property, but one has a `profile` property and the other has an `error` property.</span></span>

<span data-ttu-id="cdb98-p129">O código do manipulador na página host usa o valor da propriedade `messageType` para ramificar como no exemplo a seguir. A função `showUserName` é a mesma do exemplo anterior e a função `showNotification` exibe o erro na interface do usuário da página host.</span><span class="sxs-lookup"><span data-stu-id="cdb98-p129">The handler code in the host page uses the value of the `messageType` property to branch as shown in the following example. Note that the `showUserName` function is the same as in the previous example and `showNotification` function displays the error in the host page's UI.</span></span>

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
> <span data-ttu-id="cdb98-214">A `showNotification` implementação não é exibida no código de exemplo fornecido neste artigo.</span><span class="sxs-lookup"><span data-stu-id="cdb98-214">The `showNotification` implementation is not shown in the sample code provided by this article.</span></span> <span data-ttu-id="cdb98-215">Um exemplo de como você pode implementar essa função dentro do suplemento, confira [Exemplo do suplemento do Office exemplo do diálogo API](https://github.com/OfficeDev/Office-Add-in-Dialog-API-Simple-Example).</span><span class="sxs-lookup"><span data-stu-id="cdb98-215">For an example of how you might implement this function within your add-in, see [Office Add-in Dialog API Example](https://github.com/OfficeDev/Office-Add-in-Dialog-API-Simple-Example).</span></span>

### <a name="closing-the-dialog-box"></a><span data-ttu-id="cdb98-216">Feche a caixa de diálogo</span><span class="sxs-lookup"><span data-stu-id="cdb98-216">Closing the dialog box</span></span>

<span data-ttu-id="cdb98-p131">Você pode implementar um botão na caixa de diálogo para fechá-la. Para fazer isso, o manipulador de eventos de clique do botão deve usar `messageParent` para informar a página host em que o botão foi clicado. Apresentamos um exemplo a seguir:</span><span class="sxs-lookup"><span data-stu-id="cdb98-p131">You can implement a button in the dialog box that will close it. To do this, the click event handler for the button should use `messageParent` to tell the host page that the button has been clicked. The following is an example:</span></span>

```js
function closeButtonClick() {
    var messageObject = {messageType: "dialogClosed"};
    var jsonMessage = JSON.stringify(messageObject);
    Office.context.ui.messageParent(jsonMessage);
}
```

<span data-ttu-id="cdb98-p132">O manipulador de página host de `DialogMessageReceived` poderia chamar `dialog.close`, como neste exemplo. (Veja exemplos anteriores que mostram como o objeto dialog é inicializado.)</span><span class="sxs-lookup"><span data-stu-id="cdb98-p132">The host page handler for `DialogMessageReceived` would call `dialog.close`, as in this example. (See previous examples that show how the dialog object is initialized.)</span></span>


```js
function processMessage(arg) {
    var messageFromDialog = JSON.parse(arg.message);
    if (messageFromDialog.messageType === "dialogClosed") {
       dialog.close();
    }
}
```

<span data-ttu-id="cdb98-p133">Mesmo quando você não tem sua própria interface de usuário de diálogo de fechar, um usuário final pode fechar a caixa de diálogo escolhendo a opção **X** no canto superior direito. Essa ação aciona o evento `DialogEventReceived`. Se seu painel do host precisar saber quando isso acontece, ele deverá declarar um manipulador para esse evento. Confira a seção [Erros e eventos na janela de diálogo](#errors-and-events-in-the-dialog-window) para ver os detalhes.</span><span class="sxs-lookup"><span data-stu-id="cdb98-p133">Even when you don't have your own close dialog UI, an end user can close the dialog box by choosing the **X** in the upper-right corner. This action triggers the `DialogEventReceived` event. If your host pane needs to know when this happens, it should declare a handler for this event. See the section [Errors and events in the dialog window](#errors-and-events-in-the-dialog-window) for details.</span></span>

## <a name="handle-errors-and-events"></a><span data-ttu-id="cdb98-226">Manipular erros e eventos</span><span class="sxs-lookup"><span data-stu-id="cdb98-226">Handle errors and events</span></span>

<span data-ttu-id="cdb98-227">Seu código deve manipular duas categorias de eventos:</span><span class="sxs-lookup"><span data-stu-id="cdb98-227">Your code should handle two categories of events:</span></span>

- <span data-ttu-id="cdb98-228">Erros retornados pela chamada de `displayDialogAsync` porque não foi possível criar a caixa de diálogo.</span><span class="sxs-lookup"><span data-stu-id="cdb98-228">Errors returned by the call of `displayDialogAsync` because the dialog box cannot be created.</span></span>
- <span data-ttu-id="cdb98-229">Erros e outros eventos na janela de diálogo.</span><span class="sxs-lookup"><span data-stu-id="cdb98-229">Errors, and other events, in the dialog window.</span></span>

### <a name="errors-from-displaydialogasync"></a><span data-ttu-id="cdb98-230">Erros de displayDialogAsync</span><span class="sxs-lookup"><span data-stu-id="cdb98-230">Errors from displayDialogAsync</span></span>

<span data-ttu-id="cdb98-231">Além dos erros gerais de sistema e de plataforma, três erros são específicos para chamar `displayDialogAsync`.</span><span class="sxs-lookup"><span data-stu-id="cdb98-231">In addition to general platform and system errors, three errors are specific to calling `displayDialogAsync`.</span></span>

|<span data-ttu-id="cdb98-232">Número do código</span><span class="sxs-lookup"><span data-stu-id="cdb98-232">Code number</span></span>|<span data-ttu-id="cdb98-233">Significado</span><span class="sxs-lookup"><span data-stu-id="cdb98-233">Meaning</span></span>|
|:-----|:-----|
|<span data-ttu-id="cdb98-234">12004</span><span class="sxs-lookup"><span data-stu-id="cdb98-234">12004</span></span>|<span data-ttu-id="cdb98-p134">O domínio que a URL transmitiu para `displayDialogAsync` não é confiável. O domínio deve ser o mesmo domínio que o da página de host (incluindo o protocolo e o número de porta).</span><span class="sxs-lookup"><span data-stu-id="cdb98-p134">The domain of the URL passed to `displayDialogAsync` is not trusted. The domain must be the same domain as the host page (including protocol and port number).</span></span>|
|<span data-ttu-id="cdb98-237">12005</span><span class="sxs-lookup"><span data-stu-id="cdb98-237">12005</span></span>|<span data-ttu-id="cdb98-p135">A URL passada para `displayDialogAsync` usa o protocolo HTTP. HTTPS é necessário. (Em algumas versões do Office, a mensagem de erro retornada com 12005 é a mesma retornada para 12004.)</span><span class="sxs-lookup"><span data-stu-id="cdb98-p135">The URL passed to `displayDialogAsync` uses the HTTP protocol. HTTPS is required. (In some versions of Office, the error message returned with 12005 is the same one returned for 12004.)</span></span>|
|<span data-ttu-id="cdb98-241"><span id="12007">12007</span></span><span class="sxs-lookup"><span data-stu-id="cdb98-241"><span id="12007">12007</span></span></span>|<span data-ttu-id="cdb98-p136">Uma caixa de diálogo já está aberta na janela do host. Uma janela do host, como um painel de tarefas, só pode ter uma caixa de diálogo aberta por vez.</span><span class="sxs-lookup"><span data-stu-id="cdb98-p136">A dialog box is already opened from this host window. A host window, such as a task pane, can only have one dialog box open at a time.</span></span>|
|<span data-ttu-id="cdb98-244">12009</span><span class="sxs-lookup"><span data-stu-id="cdb98-244">12009</span></span>|<span data-ttu-id="cdb98-245">O usuário opta por ignorar a caixa de diálogo.</span><span class="sxs-lookup"><span data-stu-id="cdb98-245">The user chose to ignore the dialog box.</span></span> <span data-ttu-id="cdb98-246">Este erro pode ocorrer em versões online do Office, em que os usuários podem optar por não permitir que um suplemento apresente uma caixa de diálogo.</span><span class="sxs-lookup"><span data-stu-id="cdb98-246">This error can occur in online versions of Office, where users may choose not to allow an add-in to present a dialog.</span></span>|

<span data-ttu-id="cdb98-247">Quando `displayDialogAsync` é chamado, ele sempre transmite um objeto [AsyncResult](/javascript/api/office/office.asyncresult) para sua função de retorno de chamada.</span><span class="sxs-lookup"><span data-stu-id="cdb98-247">When `displayDialogAsync` is called, it always passes an [AsyncResult](/javascript/api/office/office.asyncresult) object to its callback function.</span></span> <span data-ttu-id="cdb98-248">Se a chamada for bem-sucedida, ou seja, a janela de diálogo for aberta, a propriedade `value` do objeto `AsyncResult` será um objeto [Dialog](/javascript/api/office/office.dialog).</span><span class="sxs-lookup"><span data-stu-id="cdb98-248">When the call is successful - that is, the dialog window is opened - the `value` property of the `AsyncResult` object is a [Dialog](/javascript/api/office/office.dialog) object.</span></span> <span data-ttu-id="cdb98-249">Um exemplo disso encontra-se na seção [Enviar informações da caixa de diálogo para a página de host](#send-information-from-the-dialog-box-to-the-host-page).</span><span class="sxs-lookup"><span data-stu-id="cdb98-249">An example of this is in the section [Send information from the dialog box to the host page](#send-information-from-the-dialog-box-to-the-host-page).</span></span> <span data-ttu-id="cdb98-250">Quando a chamada para `displayDialogAsync` falha, a janela não é criada, a propriedade `status` do objeto `AsyncResult` é definida como `Office.AsyncResultStatus.Failed` e a propriedade `error` do objeto é preenchida.</span><span class="sxs-lookup"><span data-stu-id="cdb98-250">When the call to `displayDialogAsync` fails, the window is not created, the `status` property of the `AsyncResult` object is set to `Office.AsyncResultStatus.Failed`, and the `error` property of the object is populated.</span></span> <span data-ttu-id="cdb98-251">Você deve ter sempre um retorno de chamada que testa o `status` e responde quando é um erro.</span><span class="sxs-lookup"><span data-stu-id="cdb98-251">You should always have a callback that tests the `status` and responds when it's an error.</span></span> <span data-ttu-id="cdb98-252">Para um exemplo que simplesmente relata a mensagem de erro independentemente do número do código, veja o código a seguir:</span><span class="sxs-lookup"><span data-stu-id="cdb98-252">For an example that simply reports the error message regardless of its code number, see the following code:</span></span>

```js
var dialog;
Office.context.ui.displayDialogAsync('https://myDomain/myDialog.html',
function (asyncResult) {
    if (asyncResult.status === Office.AsyncResultStatus.Failed) {
        showNotification(asyncResult.error.code = ": " + asyncResult.error.message);
    } else {
        dialog = asyncResult.value;
        dialog.addEventHandler(Office.EventType.DialogMessageReceived, processMessage);
    }
});
```

### <a name="errors-and-events-in-the-dialog-window"></a><span data-ttu-id="cdb98-253">Erros e eventos na janela de diálogo</span><span class="sxs-lookup"><span data-stu-id="cdb98-253">Errors and events in the dialog window</span></span>

<span data-ttu-id="cdb98-254">Três erros e eventos, conhecidos por seus números de códigos, na caixa de diálogo acionarão um evento `DialogEventReceived` na página host.</span><span class="sxs-lookup"><span data-stu-id="cdb98-254">Three errors and events, known by their code numbers, in the dialog box will trigger a `DialogEventReceived` event in the host page.</span></span>

|<span data-ttu-id="cdb98-255">Número do código</span><span class="sxs-lookup"><span data-stu-id="cdb98-255">Code number</span></span>|<span data-ttu-id="cdb98-256">Significado</span><span class="sxs-lookup"><span data-stu-id="cdb98-256">Meaning</span></span>|
|:-----|:-----|
|<span data-ttu-id="cdb98-257">12002</span><span class="sxs-lookup"><span data-stu-id="cdb98-257">12002</span></span>|<span data-ttu-id="cdb98-258">Uma destas opções:</span><span class="sxs-lookup"><span data-stu-id="cdb98-258">One of the following:</span></span><br> <span data-ttu-id="cdb98-259">- Não existe uma página na URL transmitida para `displayDialogAsync`.</span><span class="sxs-lookup"><span data-stu-id="cdb98-259">- No page exists at the URL that was passed to `displayDialogAsync`.</span></span><br> <span data-ttu-id="cdb98-260">- A página transmitida para `displayDialogAsync` foi carregada, mas a caixa de diálogo foi direcionada para uma página que ela não consegue localizar nem carregar ou foi direcionada para uma URL com sintaxe inválida.</span><span class="sxs-lookup"><span data-stu-id="cdb98-260">- The page that was passed to `displayDialogAsync` loaded, but the dialog box was directed to a page that it cannot find or load, or it has been directed to a URL with invalid syntax.</span></span>|
|<span data-ttu-id="cdb98-261">12003</span><span class="sxs-lookup"><span data-stu-id="cdb98-261">12003</span></span>|<span data-ttu-id="cdb98-p139">A caixa de diálogo foi direcionada para uma URL com o protocolo HTTP. HTTPS é necessário.</span><span class="sxs-lookup"><span data-stu-id="cdb98-p139">The dialog box was directed to a URL with the HTTP protocol. HTTPS is required.</span></span>|
|<span data-ttu-id="cdb98-264">12006</span><span class="sxs-lookup"><span data-stu-id="cdb98-264">12006</span></span>|<span data-ttu-id="cdb98-265">A caixa de diálogo foi fechada, geralmente pelo usuário ter escolhido o botão **X**.</span><span class="sxs-lookup"><span data-stu-id="cdb98-265">The dialog box was closed, usually because the user chooses the **X** button.</span></span>|

<span data-ttu-id="cdb98-p140">Seu código pode atribuir um manipulador para o evento `DialogEventReceived` na chamada para `displayDialogAsync`. Apresentamos um exemplo simples a seguir:</span><span class="sxs-lookup"><span data-stu-id="cdb98-p140">Your code can assign a handler for the `DialogEventReceived` event in the call to `displayDialogAsync`. The following is a simple example:</span></span>

```js
var dialog;
Office.context.ui.displayDialogAsync('https://myDomain/myDialog.html',
    function (result) {
        dialog = result.value;
        dialog.addEventHandler(Office.EventType.DialogEventReceived, processDialogEvent);
    }
);
```

<span data-ttu-id="cdb98-268">Para obter um exemplo de um manipulador para o evento `DialogEventReceived` que cria mensagens de erro personalizadas para cada código de erro, veja o exemplo a seguir:</span><span class="sxs-lookup"><span data-stu-id="cdb98-268">For an example of a handler for the `DialogEventReceived` event that creates custom error messages for each error code, see the following example:</span></span>

```js
function processDialogEvent(arg) {
    switch (arg.error) {
        case 12002:
            showNotification("The dialog box has been directed to a page that it cannot find or load, or the URL syntax is invalid.");
            break;
        case 12003:
            showNotification("The dialog box has been directed to a URL with the HTTP protocol. HTTPS is required.");            break;
        case 12006:
            showNotification("Dialog closed.");
            break;
        default:
            showNotification("Unknown error in dialog box.");
            break;
    }
}
```

<span data-ttu-id="cdb98-269">Para ver um suplemento de exemplo que manipula erros dessa forma, confira [Exemplo de API de Caixa de diálogo do Suplemento do Office](https://github.com/OfficeDev/Office-Add-in-Dialog-API-Simple-Example).</span><span class="sxs-lookup"><span data-stu-id="cdb98-269">For a sample add-in that handles errors in this way, see [Office Add-in Dialog API Example](https://github.com/OfficeDev/Office-Add-in-Dialog-API-Simple-Example).</span></span>


## <a name="pass-information-to-the-dialog-box"></a><span data-ttu-id="cdb98-270">Transmitir informações para a caixa diálogo</span><span class="sxs-lookup"><span data-stu-id="cdb98-270">Pass information to the dialog box</span></span>

<span data-ttu-id="cdb98-p141">Às vezes, a página host precisa transmitir informações para a caixa de diálogo. Você pode fazer isso de duas maneiras principais:</span><span class="sxs-lookup"><span data-stu-id="cdb98-p141">Sometimes the host page needs to pass information to the dialog box. You can do this in two primary ways:</span></span>

- <span data-ttu-id="cdb98-273">Adicionar parâmetros de consulta à URL que é transmitida para `displayDialogAsync`.</span><span class="sxs-lookup"><span data-stu-id="cdb98-273">Add query parameters to the URL that is passed to `displayDialogAsync`.</span></span>
- <span data-ttu-id="cdb98-p142">Armazenar as informações em outro local que seja acessível para a janela do host e para a caixa de diálogo. As duas janelas não compartilham um armazenamento de sessão comum, mas *se elas tiverem o mesmo domínio* (incluindo o número da porta, se houver algum), compartilharão um [Armazenamento Local](https://www.w3schools.com/html/html5_webstorage.asp) comum.\*</span><span class="sxs-lookup"><span data-stu-id="cdb98-p142">Store the information somewhere that is accessible to both the host window and dialog box. The two windows do not share a common session storage, but *if they have the same domain* (including port number, if any),  they share a common [local storage](https://www.w3schools.com/html/html5_webstorage.asp).</span></span>

> [!NOTE]
> <span data-ttu-id="cdb98-276">\* Há um bug que afetará sua estratégia de tratamento de tokens.</span><span class="sxs-lookup"><span data-stu-id="cdb98-276">\* There is a bug that will effect your strategy for token handling.</span></span> <span data-ttu-id="cdb98-277">Se o suplemento estiver sendo executado no **Office na Web** nos navegadores Safari ou Edge, o painel de tarefas e a caixa de diálogo não compartilharão o mesmo Armazenamento Local, portanto, ele não poderá ser usado para a comunicação entre eles.</span><span class="sxs-lookup"><span data-stu-id="cdb98-277">If the add-in is running in **Office on the web** in either the Safari or Edge browser, the dialog and task pane do not share the same Local Storage, so it cannot be used to communicate between them.</span></span>

### <a name="use-local-storage"></a><span data-ttu-id="cdb98-278">Usar o armazenamento local</span><span class="sxs-lookup"><span data-stu-id="cdb98-278">Use local storage</span></span>

<span data-ttu-id="cdb98-279">Para usar o armazenamento local, seu código chama o método `setItem` do objeto `window.localStorage` na página host antes da chamada `displayDialogAsync`, como no exemplo a seguir:</span><span class="sxs-lookup"><span data-stu-id="cdb98-279">To use local storage, your code calls the `setItem` method of the `window.localStorage` object in the host page before the `displayDialogAsync` call, as in the following example:</span></span>

```js
localStorage.setItem("clientID", "15963ac5-314f-4d9b-b5a1-ccb2f1aea248");
```

<span data-ttu-id="cdb98-280">O código na janela de diálogo lê o item quando necessário, como no exemplo a seguir:</span><span class="sxs-lookup"><span data-stu-id="cdb98-280">Code in the dialog window reads the item when it's needed, as in the following example:</span></span>

```js
var clientID = localStorage.getItem("clientID");
// You can also use property syntax:
// var clientID = localStorage.clientID;
```

### <a name="use-query-parameters"></a><span data-ttu-id="cdb98-281">Usar parâmetros de consulta</span><span class="sxs-lookup"><span data-stu-id="cdb98-281">Use query parameters</span></span>

<span data-ttu-id="cdb98-282">O exemplo a seguir mostra como transmitir dados com um parâmetro de consulta:</span><span class="sxs-lookup"><span data-stu-id="cdb98-282">The following example shows how to pass data with a query parameter:</span></span>

```js
Office.context.ui.displayDialogAsync('https://myAddinDomain/myDialog.html?clientID=15963ac5-314f-4d9b-b5a1-ccb2f1aea248');
```

<span data-ttu-id="cdb98-283">Para ver um exemplo que usa essa técnica, consulte [Inserir gráficos do Excel usando o Microsoft Graph em um Suplemento do PowerPoint](https://github.com/OfficeDev/PowerPoint-Add-in-Microsoft-Graph-ASPNET-InsertChart).</span><span class="sxs-lookup"><span data-stu-id="cdb98-283">For a sample that uses this technique, see [Insert Excel charts using Microsoft Graph in a PowerPoint add-in](https://github.com/OfficeDev/PowerPoint-Add-in-Microsoft-Graph-ASPNET-InsertChart).</span></span>

<span data-ttu-id="cdb98-284">O código na janela de diálogo pode analisar a URL e ler o valor do parâmetro.</span><span class="sxs-lookup"><span data-stu-id="cdb98-284">Code in your dialog window can parse the URL and read the parameter value.</span></span>

> [!NOTE]
> <span data-ttu-id="cdb98-p144">O Office adiciona automaticamente um parâmetro de consulta chamado `_host_info` à URL que é transmitida para `displayDialogAsync`. Ele é anexado após os parâmetros de consulta personalizados, se houver algum. Ele não é anexado às URLs subsequentes para as quais a caixa de diálogo navega. No futuro, a Microsoft poderá alterar o conteúdo desse valor ou removê-lo completamente para que seu código não consiga lê-lo. O mesmo valor é adicionado ao armazenamento de sessão da caixa de diálogo. Novamente, *seu código não deve ler nem gravar esse valor*.</span><span class="sxs-lookup"><span data-stu-id="cdb98-p144">Office automatically adds a query parameter called `_host_info` to the URL that is passed to `displayDialogAsync`. (It is appended after your custom query parameters, if any. It is not appended to any subsequent URLs that the dialog box navigates to.) Microsoft may change the content of this value, or remove it entirely, in the future, so your code should not read it. The same value is added to the dialog box's session storage. Again, *your code should neither read nor write to this value*.</span></span>

## <a name="use-the-dialog-apis-to-show-a-video"></a><span data-ttu-id="cdb98-290">Use APIs de Caixa de Diálogo para exibir um vídeo</span><span class="sxs-lookup"><span data-stu-id="cdb98-290">Use the Dialog APIs to show a video</span></span>

<span data-ttu-id="cdb98-291">Para mostrar um vídeo em uma caixa de diálogo:</span><span class="sxs-lookup"><span data-stu-id="cdb98-291">To show a video in a dialog box:</span></span>

1.  <span data-ttu-id="cdb98-p145">Crie uma página cujo único conteúdo seja um iframe. O atributo `src` dos pontos do iframe para um vídeo online. O protocolo da URL do vídeo deve ser HTTP**S**. Neste artigo, chamaremos esta página de "video.dialogbox.html". Veja a seguir um exemplo da marcação:</span><span class="sxs-lookup"><span data-stu-id="cdb98-p145">Create a page whose only content is an iframe. The `src` attribute of the iframe points to an online video. The protocol of the video's URL must be HTTP**S**. In this article we'll call this page "video.dialogbox.html". The following is an example of the markup:</span></span>

    ```HTML
    <iframe class="ms-firstrun-video__player"  width="640" height="360"
        src="https://www.youtube.com/embed/XVfOe5mFbAE?rel=0&autoplay=1"
        frameborder="0" allowfullscreen>
    </iframe>
    ```

2.  <span data-ttu-id="cdb98-297">A página video.dialogbox.html deve estar no mesmo domínio que a página de host.</span><span class="sxs-lookup"><span data-stu-id="cdb98-297">The video.dialogbox.html page must be in the same domain as the host page.</span></span>
3.  <span data-ttu-id="cdb98-298">Use uma chamada de `displayDialogAsync` na página host para abrir video.dialogbox.html.</span><span class="sxs-lookup"><span data-stu-id="cdb98-298">Use a call of `displayDialogAsync` in the host page to open video.dialogbox.html.</span></span>
4.  <span data-ttu-id="cdb98-p146">Se o suplemento precisar saber quando o usuário fecha a caixa de diálogo, registre um manipulador para o evento `DialogEventReceived` e manipule o evento 12006. Para mais detalhes, confira a seção [Erros e eventos na janela de diálogo](#errors-and-events-in-the-dialog-window).</span><span class="sxs-lookup"><span data-stu-id="cdb98-p146">If your add-in needs to know when the user closes the dialog box, register a handler for the `DialogEventReceived` event and handle the 12006 event. For details, see the section [Errors and events in the dialog window](#errors-and-events-in-the-dialog-window).</span></span>

<span data-ttu-id="cdb98-301">Para ver um exemplo que mostre um vídeo na caixa de diálogo, confira a [padrão de design de roteiro de vídeo](/office/dev/add-ins/design/first-run-experience-patterns#video-placemat).</span><span class="sxs-lookup"><span data-stu-id="cdb98-301">For a sample that shows a video in a dialog box, see the [video placemat design pattern](/office/dev/add-ins/design/first-run-experience-patterns#video-placemat).</span></span>

![Captura de tela de um vídeo mostrando uma caixa de diálogo de um suplemento](../images/video-placemats-dialog-open.png)

## <a name="use-the-dialog-apis-in-an-authentication-flow"></a><span data-ttu-id="cdb98-303">Use as APIs de Caixa de Diálogo em um fluxo de autenticação</span><span class="sxs-lookup"><span data-stu-id="cdb98-303">Use the Dialog APIs in an authentication flow</span></span>

<span data-ttu-id="cdb98-304">Confira[Autenticar com a API da Caixa de Diálogo do Office](auth-with-office-dialog-api.md).</span><span class="sxs-lookup"><span data-stu-id="cdb98-304">See [Authenticate with the Office Dialog API](auth-with-office-dialog-api.md).</span></span>

## <a name="using-the-office-dialog-api-with-single-page-applications-and-client-side-routing"></a><span data-ttu-id="cdb98-305">Usar a API de Caixa de diálogo do Office com aplicativos de página única e roteamento do lado do cliente</span><span class="sxs-lookup"><span data-stu-id="cdb98-305">Using the Office Dialog API with single-page applications and client-side routing</span></span>

<span data-ttu-id="cdb98-306">Se seu suplemento usa o roteamento do lado do cliente, como os aplicativos de página única geralmente fazem, você tem a opção de transmitir a URL de uma rota para o método [displayDialogAsync](/javascript/api/office/office.ui)(*o que não recomendamos*), em vez da URL de uma página HTML completa e separada.</span><span class="sxs-lookup"><span data-stu-id="cdb98-306">If your add-in uses client-side routing, as single-page applications (SPAs) typically do, you have the option to pass the URL of a route to the [displayDialogAsync](/javascript/api/office/office.ui) method (*which we recommend against*), instead of the URL of a complete and separate HTML page.</span></span>

<span data-ttu-id="cdb98-307">A caixa de diálogo está em uma nova janela com seu próprio contexto de execução.</span><span class="sxs-lookup"><span data-stu-id="cdb98-307">The dialog box is in a new window with its own execution context.</span></span> <span data-ttu-id="cdb98-308">Se você transmitir uma rota, sua página de base e todos os códigos de inicialização e bootstrapping serão executados novamente nesse novo contexto e todas as variáveis serão definidas com seus valores iniciais na caixa de diálogo.</span><span class="sxs-lookup"><span data-stu-id="cdb98-308">If you pass a route, your base page and all its initialization and bootstrapping code run again in this new context, and any variables are set to their initial values in the dialog window.</span></span> <span data-ttu-id="cdb98-309">Essa técnica baixa e Inicia uma segunda instância do seu aplicativo na janela da caixa de diálogo, o que é parcialmente contraproducente em se tratando de um SPA (aplicativo de página única).</span><span class="sxs-lookup"><span data-stu-id="cdb98-309">So this technique downloads and launches a second instance of your application in the dialog window, which partially defeats the purpose of an SPA.</span></span> <span data-ttu-id="cdb98-310">Além disso, o código que altera as variáveis na janela de diálogo não altera a versão do painel de tarefas das mesmas variáveis.</span><span class="sxs-lookup"><span data-stu-id="cdb98-310">In addition, code that changes variables in the dialog window does not change the task pane version of the same variables.</span></span> <span data-ttu-id="cdb98-311">De forma semelhante, a janela da caixa de diálogo tem seu próprio armazenamento de sessão, que não pode ser acessado a partir do código no painel de tarefas.</span><span class="sxs-lookup"><span data-stu-id="cdb98-311">Similarly, the dialog window has its own session storage, which is not accessible from code in the task pane.</span></span>

<span data-ttu-id="cdb98-312">Portanto, se você passar uma rota para o método`displayDialogAsync`, você não teria somente um SPA; você teria duas instâncias do mesmo SPA.</span><span class="sxs-lookup"><span data-stu-id="cdb98-312">So, if you passed a route to the `displayDialogAsync` method, you wouldn't really have an SPA; you'd have two instances of the same SPA.</span></span> <span data-ttu-id="cdb98-313">Além disso, a maior parte do código na instância do painel de tarefas nunca seria usada nessa instância assim como grande parte do código na instância de caixa de diálogo também nunca seria usado nessa dada instância.</span><span class="sxs-lookup"><span data-stu-id="cdb98-313">Moreover, much of the code in the task pane instance would never be used in that instance and much of the code in the dialog instance would never be used in that instance.</span></span> <span data-ttu-id="cdb98-314">Seria como ter dois SPAs no mesmo grupo.</span><span class="sxs-lookup"><span data-stu-id="cdb98-314">It would be like having two SPAs in the same bundle.</span></span> <span data-ttu-id="cdb98-315">Se o código que você deseja executar na caixa de diálogo for complexo o suficiente, talvez você queira fazer isso explicitamente; ou seja, ter dois SPAs em pastas diferentes do mesmo domínio.</span><span class="sxs-lookup"><span data-stu-id="cdb98-315">If the code that you want to run in the dialog is sufficiently complex, you might want to do this explicitly; that is, have two SPAs in different folders of the same domain.</span></span> <span data-ttu-id="cdb98-316">Mas na maioria dos cenários, apenas a lógica simples é necessária na caixa de diálogo.</span><span class="sxs-lookup"><span data-stu-id="cdb98-316">But in most scenarios, only simple logic is needed in the dialog.</span></span> <span data-ttu-id="cdb98-317">Nesses casos, o projeto será bastante simplificado simplesmente hospedando uma página HTML simples, com JavaScript incorporado ou referenciado no domínio do seu SPA.</span><span class="sxs-lookup"><span data-stu-id="cdb98-317">In such cases, your project will be greatly simplified by simply hosting a simple HTML page, with embedded or referenced JavaScript, in the domain of your SPA.</span></span> <span data-ttu-id="cdb98-318">Passe a URL da página para o método`displayDialogAsync`.</span><span class="sxs-lookup"><span data-stu-id="cdb98-318">Pass the URL of the page to the `displayDialogAsync` method.</span></span> <span data-ttu-id="cdb98-319">Isso pode significar que você está de desviando da ideia literal de um aplicativo de página única; no entanto, como observado acima, na verdade não há uma única instância de uma SPA quando você usa a caixa de diálogo.</span><span class="sxs-lookup"><span data-stu-id="cdb98-319">This might mean that you are deviating from the literal idea of a single-page app; but as noted above you don't really have a single instance of an SPA anyway when you are using the dialog.</span></span>
