---
title: Use a API de Caixa de di?logo em seus Suplementos do Office
description: ''
ms.date: 12/04/2017
ms.openlocfilehash: b026c3c5871372c52d0b44e36c01fc44a3d2bf04
ms.sourcegitcommit: c72c35e8389c47a795afbac1b2bcf98c8e216d82
ms.translationtype: HT
ms.contentlocale: pt-BR
ms.lasthandoff: 05/23/2018
---
# <a name="use-the-dialog-api-in-your-office-add-ins"></a><span data-ttu-id="92060-102">Use a API de Caixa de Di?logo em seus Suplementos do Office</span><span class="sxs-lookup"><span data-stu-id="92060-102">Use the Dialog API in your Office Add-ins</span></span>

<span data-ttu-id="92060-p101">Voc? pode usar a [API de Caixa de di?logo](https://dev.office.com/reference/add-ins/shared/officeui) para abrir caixas de di?logo no seu Suplemento do Office. Este artigo fornece orienta??es para usar a API de Caixa de di?logo em seu Suplemento do Office.</span><span class="sxs-lookup"><span data-stu-id="92060-p101">You can use the [Dialog API](https://dev.office.com/reference/add-ins/shared/officeui) to open dialog boxes in your Office Add-in. This article provides guidance for using the Dialog API in your Office Add-in.</span></span>

> [!NOTE]
> <span data-ttu-id="92060-p102">Para informa??es sobre os programas para os quais a API de Caixa de Di?logo tem suporte no momento, confira [Conjuntos de requisitos da API de Caixa de Di?logo](https://dev.office.com/reference/add-ins/requirement-sets/dialog-api-requirement-sets). Atualmente, a API de Caixa de Di?logo tem suporte para Word, Excel, PowerPoint e Outlook.</span><span class="sxs-lookup"><span data-stu-id="92060-p102">For information about where the Dialog API is currently supported, see [Dialog API requirement sets](https://dev.office.com/reference/add-ins/requirement-sets/dialog-api-requirement-sets). The Dialog API is currently supported for Word, Excel, PowerPoint, and Outlook.</span></span>

> <span data-ttu-id="92060-107">Um cen?rio prim?rio para as APIs de Caixa de Di?logo ? habilitar a autentica??o com um recurso como o Google ou o Facebook.</span><span class="sxs-lookup"><span data-stu-id="92060-107">A primary scenario for the Dialog APIs is to enable authentication with a resource such as Google or Facebook.</span></span> <span data-ttu-id="92060-108">Se o seu suplemento exigir dados sobre o usu?rio do Office ou seus recursos acess?veis atrav?s do Microsoft Graph, como o Office 365 ou o OneDrive, recomendamos que voc? use a API de logon ?nico sempre que puder.</span><span class="sxs-lookup"><span data-stu-id="92060-108">If your add-in requires data about the Office user or their resources accessible through Microsoft Graph, such as Office 365 or OneDrive, we recommend that you use the single sign-on API whenever you can.</span></span> <span data-ttu-id="92060-109">Se voc? usa as APIs para o logon ?nico, ent?o voc? n?o precisar? da API de Caixa de di?logo.</span><span class="sxs-lookup"><span data-stu-id="92060-109">If you use the APIs for single sign-on, then you will not need the Dialog API.</span></span> <span data-ttu-id="92060-110">Para mais detalhes, consulte [Habilitar o logon ?nico para Suplementos do Office](sso-in-office-add-ins.md).</span><span class="sxs-lookup"><span data-stu-id="92060-110">For details, see [Enable single sign-on for Office Add-ins](sso-in-office-add-ins.md).</span></span>

<span data-ttu-id="92060-111">Considere abrir uma caixa de di?logo em um painel de tarefas ou suplemento de conte?do ou [comando de suplemento](../design/add-in-commands.md) para fazer o seguinte:</span><span class="sxs-lookup"><span data-stu-id="92060-111">Consider opening a dialog box from a task pane or content add-in or [add-in command](../design/add-in-commands.md) to do the following:</span></span>

- <span data-ttu-id="92060-112">Exibir p?ginas de entrada que n?o podem ser abertas diretamente em um painel de tarefas.</span><span class="sxs-lookup"><span data-stu-id="92060-112">Display sign in pages that cannot be opened directly in a task pane.</span></span>
- <span data-ttu-id="92060-113">Fornecer mais espa?o na tela, ou at? uma tela inteira, para algumas tarefas no seu suplemento.</span><span class="sxs-lookup"><span data-stu-id="92060-113">Provide more screen space, or even a full screen, for some tasks in your add-in.</span></span>
- <span data-ttu-id="92060-114">Hospedar um v?deo que seria muito pequeno se fosse confinado em um painel de tarefas.</span><span class="sxs-lookup"><span data-stu-id="92060-114">Host a video that would be too small if confined to a task pane.</span></span>

> [!NOTE]
> <span data-ttu-id="92060-p104">Como a sobreposi??o de elementos de IU n?o s?o recomend?veis, evite abrir uma caixa de di?logo em um painel de tarefas a menos que seu cen?rio o obrigue a fazer isso. Ao considerar como usar a ?rea de superf?cie de um painel de tarefas, observe que pain?is de tarefas podem ter guias. Por exemplo, confira o exemplo [Suplemento do Excel JavaScriptSalesTracker](https://github.com/OfficeDev/Excel-Add-in-JavaScript-SalesTracker).</span><span class="sxs-lookup"><span data-stu-id="92060-p104">Because overlapping UI elements are discouraged, avoid opening a dialog from a task pane unless your scenario requires it. When you consider how to use the surface area of a task pane, note that task panes can be tabbed. For an example, see the [Excel Add-in JavaScript SalesTracker](https://github.com/OfficeDev/Excel-Add-in-JavaScript-SalesTracker) sample.</span></span>

<span data-ttu-id="92060-118">A imagem abaixo mostra um exemplo de uma caixa de di?logo.</span><span class="sxs-lookup"><span data-stu-id="92060-118">The following image shows an example of a dialog box.</span></span>

![Comandos de suplemento](../images/auth-o-dialog-open.png)

<span data-ttu-id="92060-p105">A caixa de di?logo sempre abre no centro da tela. O usu?rio pode mov?-la e redimension?-la. A janela ? *n?o modal*: o usu?rio pode continuar a interagir com o documento no aplicativo do Office do host e com a p?gina host no painel de tarefas, caso houver uma.</span><span class="sxs-lookup"><span data-stu-id="92060-p105">Note that the dialog box always opens in the center of the screen. The user can move and resize it. The window is *nonmodal*--a user can continue to interact with both the document in the host Office application and with the host page in the task pane, if there is one.</span></span>

## <a name="dialog-api-scenarios"></a><span data-ttu-id="92060-123">Cen?rios da API de Caixa de di?logo</span><span class="sxs-lookup"><span data-stu-id="92060-123">Dialog API scenarios</span></span>

<span data-ttu-id="92060-124">As APIs JavaScript para Office t?m suporte para os seguintes cen?rios com um objeto [Dialog](https://dev.office.com/reference/add-ins/shared/officeui.dialog) e duas fun??es no [namespace Office.context.ui](https://dev.office.com/reference/add-ins/shared/officeui).</span><span class="sxs-lookup"><span data-stu-id="92060-124">The Office JavaScript APIs support the following scenarios with a [Dialog](https://dev.office.com/reference/add-ins/shared/officeui.dialog) object and two functions in the [Office.context.ui namespace](https://dev.office.com/reference/add-ins/shared/officeui).</span></span>

### <a name="open-a-dialog-box"></a><span data-ttu-id="92060-125">Abrir uma caixa de di?logo.</span><span class="sxs-lookup"><span data-stu-id="92060-125">Open a dialog box</span></span>

<span data-ttu-id="92060-p106">Para abrir uma caixa de di?logo, seu c?digo no painel de tarefas chama o m?todo [displayDialogAsync](https://dev.office.com/reference/add-ins/shared/officeui.displaydialogasync) e transmite a ele a URL do recurso que voc? deseja abrir. Isso geralmente ? uma p?gina, mas pode ser um m?todo controlador em um aplicativo MVC, uma rota, um m?todo de servi?o Web ou qualquer outro recurso. Neste artigo, 'p?gina' ou 'site' refere-se ao recurso na caixa de di?logo. Apresentamos um exemplo de c?digo simples a seguir.</span><span class="sxs-lookup"><span data-stu-id="92060-p106">To open a dialog box, your code in the task pane calls the [displayDialogAsync](https://dev.office.com/reference/add-ins/shared/officeui.displaydialogasync) method and passes to it the URL of the resource that you want to open. This is usually a page, but it can be a controller method in an MVC application, a route, a web service method, or any other resource. In this article, 'page' or 'website' refers to the resource in the dialog. The following code is a simple example:</span></span>

```js
Office.context.ui.displayDialogAsync('https://myAddinDomain/myDialog.html');
```

> [!NOTE]
> - <span data-ttu-id="92060-p107">A URL usa o protocolo HTTP**S**. Isso ? obrigat?rio para todas as p?ginas carregadas em uma caixa di?logo, n?o apenas para a primeira p?gina carregada.</span><span class="sxs-lookup"><span data-stu-id="92060-p107">The URL uses the HTTP**S** protocol. This is mandatory for all pages loaded in a dialog box, not just the first page loaded.</span></span>
> - <span data-ttu-id="92060-p108">O dom?nio ? o mesmo que o dom?nio da p?gina host, que pode ser a p?gina em um painel de tarefas ou o [arquivo de fun??o](https://dev.office.com/reference/add-ins/manifest/functionfile) de um comando de suplemento. Isso ? necess?rio: a p?gina, o m?todo o controlador ou outro recurso que ? passado para o m?todo `displayDialogAsync` deve estar no mesmo dom?nio que a p?gina de host.</span><span class="sxs-lookup"><span data-stu-id="92060-p108">The domain is the same as the domain of the host page, which can be the page in a task pane or the [function file](https://dev.office.com/reference/add-ins/manifest/functionfile) of an add-in command. This is required: the page, controller method, or other resource that is passed to the `displayDialogAsync` method must be in the same domain as the host page.</span></span>

<span data-ttu-id="92060-p109">Ap?s o carregamento da primeira p?gina (ou de outro recurso), um usu?rio pode ir para qualquer site (ou outro recurso) que usa HTTPS. Tamb?m ? poss?vel criar a primeira p?gina para redirecionar imediatamente para outro site.</span><span class="sxs-lookup"><span data-stu-id="92060-p109">After the first page (or other resource) is loaded, a user can go to any website (or other resource) that uses HTTPS. You can also design the first page to immediately redirect to another site.</span></span>

<span data-ttu-id="92060-136">Por padr?o, a caixa de di?logo ocupar? 80% da altura e da largura na tela do dispositivo, mas voc? pode definir porcentagens diferentes. Basta transmitir um objeto de configura??o para o m?todo, como mostra o exemplo a seguir.</span><span class="sxs-lookup"><span data-stu-id="92060-136">By default, the dialog box will occupy 80% of the height and width of the device screen, but you can set different percentages by passing a configuration object to the method, as shown in the following example:</span></span>

```js
Office.context.ui.displayDialogAsync('https://myDomain/myDialog.html', {height: 30, width: 20});
```

<span data-ttu-id="92060-137">Para ver um suplemento de exemplo que faz isso, confira [Exemplo de API de Caixa de di?logo do Suplemento do Office](https://github.com/OfficeDev/Office-Add-in-Dialog-API-Simple-Example).</span><span class="sxs-lookup"><span data-stu-id="92060-137">For a sample add-in that does this, see [Office Add-in Dialog API Example](https://github.com/OfficeDev/Office-Add-in-Dialog-API-Simple-Example).</span></span>

<span data-ttu-id="92060-p110">Defina os dois valores como 100% para ter uma verdadeira experi?ncia de tela inteira. O m?ximo real ? 99,5%, e a janela ainda poder? ser movida e redimensionada.</span><span class="sxs-lookup"><span data-stu-id="92060-p110">Set both values to 100% to get what is effectively a full screen experience. (The effective maximum is 99.5%, and the window is still moveable and resizable.)</span></span>

> [!NOTE]
> <span data-ttu-id="92060-p111">Apenas uma caixa de di?logo pode ser aberta em uma janela do host. Tentar abrir outra caixa de di?logo gera um erro. Portanto, por exemplo, se um usu?rio abrir uma caixa de di?logo no painel de tarefas, ele n?o poder? abrir uma segunda caixa de di?logo em uma p?gina diferente no painel de tarefas. No entanto, quando uma caixa de di?logo ? aberta em um [comando de suplemento](../design/add-in-commands.md), o comando abre um arquivo HTML novo (mas n?o visto) sempre que ele ? selecionado. Isso cria uma nova janela do host (n?o vista) para que cada janela possa iniciar sua pr?pria caixa de di?logo. Para obter mais informa??es, confira [Erros de displayDialogAsync](#errors-from-displaydialogasync).</span><span class="sxs-lookup"><span data-stu-id="92060-p111">You can open only one dialog box from a host window. An attempt to open another dialog box generates an error. For example, if a user opens a dialog box from a task pane, she cannot open a second dialog box, from a different page in the task pane. However, when a dialog box is opened from an [add-in command](../design/add-in-commands.md), the command opens a new (but unseen) HTML file each time it is selected. This creates a new (unseen) host window, so each such window can launch its own dialog box. For more information, see [Errors from displayDialogAsync](#errors-from-displaydialogasync).</span></span>

### <a name="take-advantage-of-a-performance-option-in-office-online"></a><span data-ttu-id="92060-146">Aproveite uma op??o de desempenho no Office Online</span><span class="sxs-lookup"><span data-stu-id="92060-146">Take advantage of a performance option in Office Online</span></span>

<span data-ttu-id="92060-p112">A propriedade `displayInIframe` ? uma propriedade adicional no objeto de configura??o que voc? pode passar para `displayDialogAsync`. Quando essa propriedade for definida como `true` e o suplemento estiver em execu??o em um documento aberto no Office Online, a caixa de di?logo ser? aberta como um iframe flutuante, em vez de uma janela independente, o que faz com que ela seja aberta mais rapidamente. Apresentamos um exemplo a seguir:</span><span class="sxs-lookup"><span data-stu-id="92060-p112">The `displayInIframe` property is an additional property in the configuration object that you can pass to `displayDialogAsync`. When this property is set to `true`, and the add-in is running in a document opened in Office Online, the dialog box will open as a floating iframe rather than an independent window, which makes it open faster. The following is an example:</span></span>

```js
Office.context.ui.displayDialogAsync('https://myDomain/myDialog.html', {height: 30, width: 20, displayInIframe: true});
```

<span data-ttu-id="92060-150">O valor padr?o ? `false`, que ? o mesmo que omitir a propriedade inteiramente.</span><span class="sxs-lookup"><span data-stu-id="92060-150">The default value is `false`, which is the same as omitting the property entirely.</span></span> <span data-ttu-id="92060-151">Se o suplemento n?o estiver sendo executado no Office Online, `displayInIframe` ser? ignorado.</span><span class="sxs-lookup"><span data-stu-id="92060-151">If the add-in is not running in Office Online, the `displayInIframe` is ignored.</span></span>

> [!NOTE]
> <span data-ttu-id="92060-p114">Voc? **n?o** dever? usar `displayInIframe: true` se a caixa de di?logo redirecionar a qualquer ponto para uma p?gina que n?o possa ser aberta em um iframe. Por exemplo, as p?ginas de entrada de muitos servi?os Web populares, como Google e Conta da Microsoft, n?o podem ser abertas em um iframe.</span><span class="sxs-lookup"><span data-stu-id="92060-p114">You should **not** use `displayInIframe: true` if the dialog will at any point redirect to a page that cannot be opened in an iframe. For example, the sign in pages of many popular web services, such as Google and Microsoft Account, cannot be opened in an iframe.</span></span>

### <a name="send-information-from-the-dialog-box-to-the-host-page"></a><span data-ttu-id="92060-154">Envie informa??es da caixa de di?logo para a p?gina host</span><span class="sxs-lookup"><span data-stu-id="92060-154">Send information from the dialog box to the host page</span></span>

<span data-ttu-id="92060-155">A caixa de di?logo n?o pode se comunicar com a p?gina host no painel de tarefas, a menos que:</span><span class="sxs-lookup"><span data-stu-id="92060-155">The dialog box cannot communicate with the host page in the task pane unless:</span></span>

- <span data-ttu-id="92060-156">A p?gina atual na caixa de di?logo esteja no mesmo dom?nio da p?gina host.</span><span class="sxs-lookup"><span data-stu-id="92060-156">The current page in the dialog box is in the same domain as the host page.</span></span>
- <span data-ttu-id="92060-p115">A biblioteca JavaScript do Office seja carregada na p?gina. Como qualquer p?gina que usa a biblioteca JavaScript do Office, o script da p?gina deve atribuir um m?todo ? propriedade `Office.initialize`, embora ele possa ser um m?todo vazio. Para mais detalhes, confira [Iniciar o suplemento](understanding-the-javascript-api-for-office.md#initializing-your-add-in).</span><span class="sxs-lookup"><span data-stu-id="92060-p115">The Office JavaScript library is loaded in the page. (Like any page that uses the Office JavaScript library, script for the page must assign a method to the `Office.initialize` property, although it can be an empty method. For details, see [Initializing your add-in](understanding-the-javascript-api-for-office.md#initializing-your-add-in).)</span></span>

<span data-ttu-id="92060-p116">O c?digo na p?gina de di?logo use a fun??o `messageParent` para enviar uma mensagem de cadeia de caracteres ou um valor booliano para a p?gina host. A cadeia de caracteres pode ser uma palavra, uma frase, um blob XML, um JSON em formato de cadeia de caracteres ou qualquer outra coisa que possa ser serializada em uma cadeia de caracteres. Apresentamos um exemplo a seguir:</span><span class="sxs-lookup"><span data-stu-id="92060-p116">Code in the dialog page uses the `messageParent` function to send either a Boolean value or a string message to the host page. The string can be a word, sentence, XML blob, stringified JSON, or anything else that can be serialized to a string. The following is an example:</span></span>

```js
if (loginSuccess) {
    Office.context.ui.messageParent(true);
}
```

> [!NOTE]
> - <span data-ttu-id="92060-p117">A fun??o `messageParent` ? uma das *?nicas* duas APIs do Office que pode ser chamada na caixa de di?logo. A outra ? `Office.context.requirements.isSetSupported`. Para saber mais, confira [Especificar hosts do Office e requisitos da API](specify-office-hosts-and-api-requirements.md).</span><span class="sxs-lookup"><span data-stu-id="92060-p117">The `messageParent` function is one of *only* two Office APIs that can be called in the dialog box. The other is `Office.context.requirements.isSetSupported`. For information about it, see [Specify Office hosts and API requirements](specify-office-hosts-and-api-requirements.md).</span></span>
> - <span data-ttu-id="92060-166">A fun??o `messageParent` s? pode ser chamada em uma p?gina com o mesmo dom?nio (incluindo o protocolo e a porta) da p?gina host.</span><span class="sxs-lookup"><span data-stu-id="92060-166">The `messageParent` function can only be called on a page with the same domain (including protocol and port) as the host page.</span></span>

<span data-ttu-id="92060-167">No pr?ximo exemplo, `googleProfile` ? uma vers?o em formato de cadeia de caracteres do perfil do Google do usu?rio.</span><span class="sxs-lookup"><span data-stu-id="92060-167">In the next example, `googleProfile` is a stringified version of the user's Google profile.</span></span>

```js
if (loginSuccess) {
    Office.context.ui.messageParent(googleProfile);
}
```

<span data-ttu-id="92060-p118">A p?gina host deve ser configurada para receber a mensagem. Voc? pode fazer isso adicionando um par?metro de retorno de chamada ? chamada original de `displayDialogAsync`. O retorno de chamada atribui um manipulador ao evento `DialogMessageReceived`. Apresentamos um exemplo a seguir:</span><span class="sxs-lookup"><span data-stu-id="92060-p118">The host page must be configured to receive the message. You do this by adding a callback parameter to the original call of `displayDialogAsync`. The callback assigns a handler to the `DialogMessageReceived` event. The following is an example:</span></span>

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
> - <span data-ttu-id="92060-p119">O Office transmite um objeto [AsyncResult](https://dev.office.com/reference/add-ins/shared/asyncresult) para o retorno de chamada. Ele representa o resultado de tentativas de abrir a caixa de di?logo, mas n?o representa o resultado de eventos na caixa di?logo. Para obter mais informa??es sobre essa distin??o, confira a se??o [Manipular erros e eventos](#handle-errors-and-events).</span><span class="sxs-lookup"><span data-stu-id="92060-p119">Office passes an [AsyncResult](https://dev.office.com/reference/add-ins/shared/asyncresult) object to the callback. It represents the result of the attempt to open the dialog box. It does not represent the outcome of any events in the dialog box. For more on this distinction, see the section [Handle errors and events](#handle-errors-and-events).</span></span>
> - <span data-ttu-id="92060-176">A propriedade `value` do `asyncResult` ? definida como um objeto [Dialog](https://dev.office.com/reference/add-ins/shared/officeui.dialog) que existe na p?gina host, n?o no contexto da execu??o da caixa de di?logo.</span><span class="sxs-lookup"><span data-stu-id="92060-176">The `value` property of the `asyncResult` is set to a [Dialog](https://dev.office.com/reference/add-ins/shared/officeui.dialog) object, which exists in the host page, not in the dialog box's execution context.</span></span>
> - <span data-ttu-id="92060-p120">O `processMessage` ? a fun??o que manipula o evento. Voc? pode dar a ele o nome que desejar.</span><span class="sxs-lookup"><span data-stu-id="92060-p120">The `processMessage` is the function that handles the event. You can give it any name you want.</span></span>
> - <span data-ttu-id="92060-179">A vari?vel `dialog` ? declarada em um escopo mais amplo do que o retorno de chamada porque ela tamb?m ? referenciada em `processMessage`.</span><span class="sxs-lookup"><span data-stu-id="92060-179">The `dialog` variable is declared at a wider scope than the callback because it is also referenced in `processMessage`.</span></span>

<span data-ttu-id="92060-180">Veja a seguir um exemplo simples de um manipulador para o evento `DialogMessageReceived`:</span><span class="sxs-lookup"><span data-stu-id="92060-180">The following is a simple example of a handler for the `DialogMessageReceived` event:</span></span>

```js
function processMessage(arg) {
    var messageFromDialog = JSON.parse(arg.message);
    showUserName(messageFromDialog.name);
}
```

> [!NOTE]
> - <span data-ttu-id="92060-p121">O Office transmite o objeto `arg` para o manipulador. Sua propriedade `message` ? o booliano ou a cadeia de caracteres enviada pela chamada de `messageParent` na caixa de di?logo. Neste exemplo, ela ? uma representa??o em formato de cadeia de caracteres de um perfil de usu?rio de um servi?o como a Conta da Microsoft ou o Google, portanto est? desserializada como um objeto com `JSON.parse` novamente.</span><span class="sxs-lookup"><span data-stu-id="92060-p121">Office passes the `arg` object to the handler. Its `message` property is the Boolean or string sent by the call of `messageParent` in the dialog. In this example, it is a stringified representation of a user's profile from a service such as Microsoft Account or Google, so it is deserialized back to an object with `JSON.parse`.</span></span>
> - <span data-ttu-id="92060-p122">A implementa??o de `showUserName` n?o ? mostrada. Ela pode exibir uma mensagem de boas-vindas personalizada no painel de tarefas.</span><span class="sxs-lookup"><span data-stu-id="92060-p122">The `showUserName` implementation is not shown. It might display a personalized welcome message on the task pane.</span></span>

<span data-ttu-id="92060-186">Quando a intera??o do usu?rio com a caixa de di?logo for conclu?da, seu manipulador de mensagem fechar? a caixa de di?logo, conforme mostrado neste exemplo.</span><span class="sxs-lookup"><span data-stu-id="92060-186">When the user interaction with the dialog box is completed, your message handler should close the dialog box, as shown in this example.</span></span>

```js
function processMessage(arg) {
    dialog.close();
    // message processing code goes here;
}
```

> [!NOTE]
> - <span data-ttu-id="92060-187">O objeto `dialog` deve ser o mesmo que ? retornado pela chamada de `displayDialogAsync`.</span><span class="sxs-lookup"><span data-stu-id="92060-187">The `dialog` object must be the same one that is returned by the call of `displayDialogAsync`.</span></span>
> - <span data-ttu-id="92060-188">A chamada de `dialog.close` informa ao Office para fechar a caixa de di?logo imediatamente.</span><span class="sxs-lookup"><span data-stu-id="92060-188">The call of `dialog.close` tells Office to immediately close the dialog box.</span></span>

<span data-ttu-id="92060-189">Para ver um suplemento de exemplo que usa essas t?cnicas, confira [Exemplo de API de Caixa de di?logo do Suplemento do Office](https://github.com/OfficeDev/Office-Add-in-Dialog-API-Simple-Example).</span><span class="sxs-lookup"><span data-stu-id="92060-189">For a sample add-in that uses these techniques, see [Office Add-in Dialog API Example](https://github.com/OfficeDev/Office-Add-in-Dialog-API-Simple-Example).</span></span>

<span data-ttu-id="92060-p123">Se o suplemento precisa abrir uma p?gina diferente do painel de tarefas depois de receber a mensagem, ? poss?vel usar o m?todo `window.location.replace` (ou `window.location.href`) como a ?ltima linha do manipulador. Apresentamos um exemplo a seguir:</span><span class="sxs-lookup"><span data-stu-id="92060-p123">If the add-in needs to open a different page of the task pane after receiving the message, you can use the `window.location.replace` method (or `window.location.href`) as the last line of the handler. The following is an example:</span></span>

```js
function processMessage(arg) {
    // message processing code goes here;
    window.location.replace("/newPage.html");
    // Alternatively ...
    // window.location.href = "/newPage.html";
}
```

<span data-ttu-id="92060-192">Para ver um exemplo de um suplemento que faz isso, confira o exemplo [Inserir gr?ficos do Excel usando o Microsoft Graph em um Suplemento do PowerPoint](https://github.com/OfficeDev/PowerPoint-Add-in-Microsoft-Graph-ASPNET-InsertChart).</span><span class="sxs-lookup"><span data-stu-id="92060-192">For an example of an add-in that does this, see the [Insert Excel charts using Microsoft Graph in a PowerPoint Add-in](https://github.com/OfficeDev/PowerPoint-Add-in-Microsoft-Graph-ASPNET-InsertChart) sample.</span></span>

#### <a name="conditional-messaging"></a><span data-ttu-id="92060-193">Mensagens condicionais</span><span class="sxs-lookup"><span data-stu-id="92060-193">Conditional messaging</span></span>
<span data-ttu-id="92060-p124">Como voc? pode enviar v?rias chamadas `messageParent` a partir da caixa de di?logo, mas tem apenas um manipulador na p?gina host do evento `DialogMessageReceived`, o manipulador tem que usar a l?gica condicional para distinguir mensagens diferentes. Por exemplo, se a caixa de di?logo solicitar que o usu?rio entre em um provedor de identidade como a Conta da Microsoft ou o Google, ele enviar? o perfil do usu?rio como uma mensagem. Se a autentica??o falhar, a caixa de di?logo enviar? informa??es de erro ? p?gina host, como no exemplo a seguir:</span><span class="sxs-lookup"><span data-stu-id="92060-p124">Because you can send multiple `messageParent` calls from the dialog box, but you have only one handler in the host page for the `DialogMessageReceived` event, the handler must use conditional logic to distinguish different messages. For example, if the dialog box prompts a user to sign in to an identity provider such as Microsoft Account or Google, it sends the user's profile as a message. If authentication fails, the dialog box sends error information to the host page, as in the following example:</span></span>

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
> - <span data-ttu-id="92060-197">A vari?vel `loginSuccess` poderia ser inicializada por meio da leitura da resposta HTTP no provedor de identidade.</span><span class="sxs-lookup"><span data-stu-id="92060-197">The `loginSuccess` variable would be initialized by reading the HTTP response from the identity provider.</span></span>
> - <span data-ttu-id="92060-p125">A implementa??o das fun??es `getProfile` e `getError` n?o ? exibida. Cada uma delas obt?m dados de um par?metro de consulta ou do corpo da resposta HTTP.</span><span class="sxs-lookup"><span data-stu-id="92060-p125">The the implementation of the `getProfile` and `getError` functions are not not shown. They each get data from a query parameter or from the body of the HTTP response.</span></span>
> - <span data-ttu-id="92060-p126">S?o enviados objetos an?nimos de diferentes tipos se a entrada for bem-sucedida ou n?o. Ambos t?m uma propriedade `messageType`, mas um tem uma propriedade `profile` e o outro tem uma propriedade `error`.</span><span class="sxs-lookup"><span data-stu-id="92060-p126">Anonymous objects of different types are sent depending on whether the sign in was successful. Both have a `messageType` property, but one has a `profile` property and the other has an `error` property.</span></span>

<span data-ttu-id="92060-202">Para ver exemplos que usam mensagens condicionais, confira</span><span class="sxs-lookup"><span data-stu-id="92060-202">For samples that use conditional messaging, see:</span></span>
- [<span data-ttu-id="92060-203">Suplemento do Office que usa o servi?o Auth0 para facilitar o login social</span><span class="sxs-lookup"><span data-stu-id="92060-203">Office Add-in that uses the Auth0 Service to Simplify Social Login</span></span>](https://github.com/OfficeDev/Office-Add-in-Auth0)
- [<span data-ttu-id="92060-204">Suplemento do Office que usa o Servi?o do OAuth.io para Simplificar o Acesso a Servi?os Populares Online</span><span class="sxs-lookup"><span data-stu-id="92060-204">Office Add-in that uses the OAuth.io Service to Simplify Access to Popular Online Services</span></span>](https://github.com/OfficeDev/Office-Add-in-OAuth.io)

<span data-ttu-id="92060-p127">O c?digo do manipulador na p?gina host usa o valor da propriedade `messageType` para ramificar como no exemplo a seguir. A fun??o `showUserName` ? a mesma do exemplo anterior e a fun??o `showNotification` exibe o erro na interface do usu?rio da p?gina host.</span><span class="sxs-lookup"><span data-stu-id="92060-p127">The handler code in the host page uses the value of the `messageType` property to branch as shown in the following example. Note that the `showUserName` function is the same as in the previous example and `showNotification` function displays the error in the host page's UI.</span></span>

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

### <a name="closing-the-dialog-box"></a><span data-ttu-id="92060-207">Feche a caixa de di?logo</span><span class="sxs-lookup"><span data-stu-id="92060-207">Closing the dialog box</span></span>

<span data-ttu-id="92060-p128">Voc? pode implementar um bot?o na caixa de di?logo para fech?-la. Para fazer isso, o manipulador de eventos de clique do bot?o deve usar `messageParent` para informar a p?gina host em que o bot?o foi clicado. Apresentamos um exemplo a seguir:</span><span class="sxs-lookup"><span data-stu-id="92060-p128">You can implement a button in the dialog box that will close it. To do this, the click event handler for the button should use `messageParent` to tell the host page that the button has been clicked. The following is an example:</span></span>

```js
function closeButtonClick() {
    var messageObject = {messageType: "dialogClosed"};            
    var jsonMessage = JSON.stringify(messageObject);
    Office.context.ui.messageParent(jsonMessage);
}
```

<span data-ttu-id="92060-p129">O manipulador de p?gina host de `DialogMessageReceived` poderia chamar `dialog.close`, como neste exemplo. (Veja exemplos anteriores que mostram como o objeto dialog ? inicializado.)</span><span class="sxs-lookup"><span data-stu-id="92060-p129">The host page handler for `DialogMessageReceived` would call `dialog.close`, as in this example. (See previous examples that show how the dialog object is initialized.)</span></span>


```js
function processMessage(arg) {
    var messageFromDialog = JSON.parse(arg.message);
    if (messageFromDialog.messageType === "dialogClosed") {
       dialog.close();
    }
}
```

<span data-ttu-id="92060-213">Para ver um exemplo que usa essa t?cnica, confira o [padr?o de design da navega??o do di?logo](https://github.com/OfficeDev/Office-Add-in-UX-Design-Patterns-Code/tree/master/templates/dialog/navigation) no reposit?rio [padr?es de design da experi?ncia do usu?rio para Suplementos do Office](https://github.com/OfficeDev/Office-Add-in-UX-Design-Patterns-Code).</span><span class="sxs-lookup"><span data-stu-id="92060-213">For a sample that uses this technique, see the [dialog navigation design pattern](https://github.com/OfficeDev/Office-Add-in-UX-Design-Patterns-Code/tree/master/templates/dialog/navigation) in the [UX design patterns for Office Add-ins](https://github.com/OfficeDev/Office-Add-in-UX-Design-Patterns-Code) repo.</span></span>

<span data-ttu-id="92060-p130">Mesmo quando voc? n?o tem sua pr?pria interface de usu?rio de di?logo de fechar, um usu?rio final pode fechar a caixa de di?logo escolhendo a op??o **X** no canto superior direito. Essa a??o aciona o evento `DialogEventReceived`. Se seu painel do host precisar saber quando isso acontece, ele dever? declarar um manipulador para esse evento. Confira a se??o [Erros e eventos na janela de di?logo](#errors-and-events-in-the-dialog-window) para ver os detalhes.</span><span class="sxs-lookup"><span data-stu-id="92060-p130">Even when you don't have your own close dialog UI, an end user can close the dialog box by choosing the **X** in the upper-right corner. This action triggers the `DialogEventReceived` event. If your host pane needs to know when this happens, it should declare a handler for this event. See the section [Errors and events in the dialog window](#errors-and-events-in-the-dialog-window) for details.</span></span>

## <a name="handle-errors-and-events"></a><span data-ttu-id="92060-218">Manipular erros e eventos</span><span class="sxs-lookup"><span data-stu-id="92060-218">Handle errors and events</span></span>

<span data-ttu-id="92060-219">Seu c?digo deve manipular duas categorias de eventos:</span><span class="sxs-lookup"><span data-stu-id="92060-219">Your code should handle two categories of events:</span></span>

- <span data-ttu-id="92060-220">Erros retornados pela chamada de `displayDialogAsync` porque n?o foi poss?vel criar a caixa de di?logo.</span><span class="sxs-lookup"><span data-stu-id="92060-220">Errors returned by the call of `displayDialogAsync` because the dialog box cannot be created.</span></span>
- <span data-ttu-id="92060-221">Erros e outros eventos na janela de di?logo.</span><span class="sxs-lookup"><span data-stu-id="92060-221">Errors, and other events, in the dialog window.</span></span>

### <a name="errors-from-displaydialogasync"></a><span data-ttu-id="92060-222">Erros de displayDialogAsync</span><span class="sxs-lookup"><span data-stu-id="92060-222">Errors from displayDialogAsync</span></span>

<span data-ttu-id="92060-223">Al?m dos erros gerais de sistema e de plataforma, tr?s erros s?o espec?ficos para chamar `displayDialogAsync`.</span><span class="sxs-lookup"><span data-stu-id="92060-223">In addition to general platform and system errors, three errors are specific to calling `displayDialogAsync`.</span></span>

|<span data-ttu-id="92060-224">N?mero do c?digo</span><span class="sxs-lookup"><span data-stu-id="92060-224">Code number</span></span>|<span data-ttu-id="92060-225">Significado</span><span class="sxs-lookup"><span data-stu-id="92060-225">Meaning</span></span>|
|:-----|:-----|
|<span data-ttu-id="92060-226">12004</span><span class="sxs-lookup"><span data-stu-id="92060-226">12004</span></span>|<span data-ttu-id="92060-p131">O dom?nio que a URL transmitiu para `displayDialogAsync` n?o ? confi?vel. O dom?nio deve ser o mesmo dom?nio que o da p?gina de host (incluindo o protocolo e o n?mero de porta).</span><span class="sxs-lookup"><span data-stu-id="92060-p131">The domain of the URL passed to `displayDialogAsync` is not trusted. The domain must be the same domain as the host page (including protocol and port number).</span></span>|
|<span data-ttu-id="92060-229">12005</span><span class="sxs-lookup"><span data-stu-id="92060-229">12005</span></span>|<span data-ttu-id="92060-p132">A URL passada para `displayDialogAsync` usa o protocolo HTTP. HTTPS ? necess?rio. (Em algumas vers?es do Office, a mensagem de erro retornada com 12005 ? a mesma retornada para 12004.)</span><span class="sxs-lookup"><span data-stu-id="92060-p132">The URL passed to `displayDialogAsync` uses the HTTP protocol. HTTPS is required. (In some versions of Office, the error message returned with 12005 is the same one returned for 12004.)</span></span>|
|<span data-ttu-id="92060-233"><span id="12007">12007</span></span><span class="sxs-lookup"><span data-stu-id="92060-233"><span id="12007">12007</span></span></span>|<span data-ttu-id="92060-p133">Uma caixa de di?logo j? est? aberta na janela do host. Uma janela do host, como um painel de tarefas, s? pode ter uma caixa de di?logo aberta por vez.</span><span class="sxs-lookup"><span data-stu-id="92060-p133">A dialog box is already opened from this host window. A host window, such as a task pane, can only have one dialog box open at a time.</span></span>|

<span data-ttu-id="92060-p134">Quando `displayDialogAsync` ? chamado, ele sempre transmite um objeto [AsyncResult](https://dev.office.com/reference/add-ins/shared/asyncresult) para sua fun??o de retorno de chamada. Se a chamada for bem-sucedida, ou seja, a janela de di?logo for aberta, a propriedade `value` do objeto `AsyncResult` ser? um objeto [Dialog](https://dev.office.com/reference/add-ins/shared/officeui.dialog). Um exemplo disso encontra-se na se??o [Enviar informa??es da caixa de di?logo para a p?gina host](#send-information-from-the-dialog-box-to-the-host-page). Quando a chamada para `displayDialogAsync` falha, a janela n?o ? criada, a propriedade `status` do objeto `AsyncResult` ? definida como "falha" e a propriedade `error` do objeto ? preenchida. Voc? deve ter sempre um retorno de chamada que testa o `status` e responde quando ? um erro. Veja a seguir um exemplo que simplesmente relata a mensagem de erro independentemente do n?mero do c?digo:</span><span class="sxs-lookup"><span data-stu-id="92060-p134">When `displayDialogAsync` is called, it always passes an [AsyncResult](https://dev.office.com/reference/add-ins/shared/asyncresult) object to its callback function. When the call is successful - that is, the dialog window is opened - the `value` property of the `AsyncResult` object is a [Dialog](https://dev.office.com/reference/add-ins/shared/officeui.dialog) object. An example of this is in the section [Send information from the dialog box to the host page](#send-information-from-the-dialog-box-to-the-host-page). When the call to `displayDialogAsync` fails, the window is not created, the `status` property of the `AsyncResult` object is set to "failed", and the `error` property of the object is populated. You should always have a callback that tests the `status` and responds when it's an error. For an example that simply reports the error message regardless of its code number, see the following code:</span></span>

```js
var dialog;
Office.context.ui.displayDialogAsync('https://myDomain/myDialog.html',
function (asyncResult) {
    if (asyncResult.status === "failed") {
        showNotification(asynceResult.error.code = ": " + asyncResult.error.message);
    } else {
        dialog = asyncResult.value;
        dialog.addEventHandler(Office.EventType.DialogMessageReceived, processMessage);
    }
});
```

### <a name="errors-and-events-in-the-dialog-window"></a><span data-ttu-id="92060-242">Erros e eventos na janela de di?logo</span><span class="sxs-lookup"><span data-stu-id="92060-242">Errors and events in the dialog window</span></span>

<span data-ttu-id="92060-243">Tr?s erros e eventos, conhecidos por seus n?meros de c?digos, na caixa de di?logo acionar?o um evento `DialogEventReceived` na p?gina host.</span><span class="sxs-lookup"><span data-stu-id="92060-243">Three errors and events, known by their code numbers, in the dialog box will trigger a `DialogEventReceived` event in the host page.</span></span>

|<span data-ttu-id="92060-244">N?mero do c?digo</span><span class="sxs-lookup"><span data-stu-id="92060-244">Code number</span></span>|<span data-ttu-id="92060-245">Significado</span><span class="sxs-lookup"><span data-stu-id="92060-245">Meaning</span></span>|
|:-----|:-----|
|<span data-ttu-id="92060-246">12002</span><span class="sxs-lookup"><span data-stu-id="92060-246">12002</span></span>|<span data-ttu-id="92060-247">Uma destas op??es:</span><span class="sxs-lookup"><span data-stu-id="92060-247">One of the following:</span></span><br> <span data-ttu-id="92060-248">- N?o existe uma p?gina na URL transmitida para `displayDialogAsync`.</span><span class="sxs-lookup"><span data-stu-id="92060-248">- No page exists at the URL that was passed to `displayDialogAsync`.</span></span><br> <span data-ttu-id="92060-249">- A p?gina transmitida para `displayDialogAsync` foi carregada, mas a caixa de di?logo foi direcionada para uma p?gina que ela n?o consegue localizar nem carregar ou foi direcionada para uma URL com sintaxe inv?lida.</span><span class="sxs-lookup"><span data-stu-id="92060-249">- The page that was passed to `displayDialogAsync` loaded, but the dialog box was directed to a page that it cannot find or load, or it has been directed to a URL with invalid syntax.</span></span>|
|<span data-ttu-id="92060-250">12003</span><span class="sxs-lookup"><span data-stu-id="92060-250">12003</span></span>|<span data-ttu-id="92060-p135">A caixa de di?logo foi direcionada para uma URL com o protocolo HTTP. HTTPS ? necess?rio.</span><span class="sxs-lookup"><span data-stu-id="92060-p135">The dialog box was directed to a URL with the HTTP protocol. HTTPS is required.</span></span>|
|<span data-ttu-id="92060-253">12006</span><span class="sxs-lookup"><span data-stu-id="92060-253">12006</span></span>|<span data-ttu-id="92060-254">A caixa de di?logo foi fechada, geralmente pelo usu?rio ter escolhido o bot?o **X**.</span><span class="sxs-lookup"><span data-stu-id="92060-254">The dialog box was closed, usually because the user chooses the **X** button.</span></span>|

<span data-ttu-id="92060-p136">Seu c?digo pode atribuir um manipulador para o evento `DialogEventReceived` na chamada para `displayDialogAsync`. Apresentamos um exemplo simples a seguir:</span><span class="sxs-lookup"><span data-stu-id="92060-p136">Your code can assign a handler for the `DialogEventReceived` event in the call to `displayDialogAsync`. The following is a simple example:</span></span>

```js
var dialog;
Office.context.ui.displayDialogAsync('https://myDomain/myDialog.html',
    function (result) {
        dialog = result.value;
        dialog.addEventHandler(Office.EventType.DialogEventReceived, processDialogEvent);
    }
);
```

<span data-ttu-id="92060-257">Para obter um exemplo de um manipulador para o evento `DialogEventReceived` que cria mensagens de erro personalizadas para cada c?digo de erro, veja o exemplo a seguir:</span><span class="sxs-lookup"><span data-stu-id="92060-257">For an example of a handler for the `DialogEventReceived` event that creates custom error messages for each error code, see the following example:</span></span>

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

<span data-ttu-id="92060-258">Para ver um suplemento de exemplo que manipula erros dessa forma, confira [Exemplo de API de Caixa de di?logo do Suplemento do Office](https://github.com/OfficeDev/Office-Add-in-Dialog-API-Simple-Example).</span><span class="sxs-lookup"><span data-stu-id="92060-258">For a sample add-in that handles errors in this way, see [Office Add-in Dialog API Example](https://github.com/OfficeDev/Office-Add-in-Dialog-API-Simple-Example).</span></span>


## <a name="pass-information-to-the-dialog-box"></a><span data-ttu-id="92060-259">Transmitir informa??es para a caixa di?logo</span><span class="sxs-lookup"><span data-stu-id="92060-259">Pass information to the dialog box</span></span>

<span data-ttu-id="92060-p137">?s vezes, a p?gina host precisa transmitir informa??es para a caixa de di?logo. Voc? pode fazer isso de duas maneiras principais:</span><span class="sxs-lookup"><span data-stu-id="92060-p137">Sometimes the host page needs to pass information to the dialog box. You can do this in two primary ways:</span></span>

- <span data-ttu-id="92060-262">Adicionar par?metros de consulta ? URL que ? transmitida para `displayDialogAsync`.</span><span class="sxs-lookup"><span data-stu-id="92060-262">Add query parameters to the URL that is passed to `displayDialogAsync`.</span></span>
- <span data-ttu-id="92060-p138">Armazenar as informa??es em outro local que seja acess?vel para a janela do host e para a caixa de di?logo. As duas janelas n?o compartilham um armazenamento de sess?o comum, mas *se elas tiverem o mesmo dom?nio* (incluindo o n?mero da porta, se houver algum), compartilhar?o um [local de armazenamento](http://www.w3schools.com/html/html5_webstorage.asp) comum.</span><span class="sxs-lookup"><span data-stu-id="92060-p138">Store the information somewhere that is accessible to both the host window and dialog box. The two windows do not share a common session storage, but *if they have the same domain* (including port number, if any),  they share a common [local storage](http://www.w3schools.com/html/html5_webstorage.asp).</span></span>

### <a name="use-local-storage"></a><span data-ttu-id="92060-265">Usar o armazenamento local</span><span class="sxs-lookup"><span data-stu-id="92060-265">Use local storage</span></span>

<span data-ttu-id="92060-266">Para usar o armazenamento local, seu c?digo chama o m?todo `setItem` do objeto `window.localStorage` na p?gina host antes da chamada `displayDialogAsync`, como no exemplo a seguir:</span><span class="sxs-lookup"><span data-stu-id="92060-266">To use local storage, your code calls the `setItem` method of the `window.localStorage` object in the host page before the `displayDialogAsync` call, as in the following example:</span></span>

```js
localStorage.setItem("clientID", "15963ac5-314f-4d9b-b5a1-ccb2f1aea248");
```

<span data-ttu-id="92060-267">O c?digo na janela de di?logo l? o item quando necess?rio, como no exemplo a seguir:</span><span class="sxs-lookup"><span data-stu-id="92060-267">Code in the dialog window reads the item when it's needed, as in the following example:</span></span>

```js
var clientID = localStorage.getItem("clientID");
// You can also use property syntax:
// var clientID = localStorage.clientID;
```

<span data-ttu-id="92060-268">Para ver exemplos de suplementos que usam o armazenamento local dessa forma, confira:</span><span class="sxs-lookup"><span data-stu-id="92060-268">For sample add-ins that uses local storage in this way, see:</span></span>

- [<span data-ttu-id="92060-269">Suplemento do Office que usa o servi?o Auth0 para facilitar o login social</span><span class="sxs-lookup"><span data-stu-id="92060-269">Office Add-in that uses the Auth0 Service to Simplify Social Login</span></span>](https://github.com/OfficeDev/Office-Add-in-Auth0)
- [<span data-ttu-id="92060-270">Suplemento do Office que usa o Servi?o do OAuth.io para Simplificar o Acesso a Servi?os Populares Online</span><span class="sxs-lookup"><span data-stu-id="92060-270">Office Add-in that uses the OAuth.io Service to Simplify Access to Popular Online Services</span></span>](https://github.com/OfficeDev/Office-Add-in-OAuth.io)

### <a name="use-query-parameters"></a><span data-ttu-id="92060-271">Usar par?metros de consulta</span><span class="sxs-lookup"><span data-stu-id="92060-271">Use query parameters</span></span>

<span data-ttu-id="92060-272">O exemplo a seguir mostra como transmitir dados com um par?metro de consulta:</span><span class="sxs-lookup"><span data-stu-id="92060-272">The following example shows how to pass data with a query parameter:</span></span>

```js
Office.context.ui.displayDialogAsync('https://myAddinDomain/myDialog.html?clientID=15963ac5-314f-4d9b-b5a1-ccb2f1aea248');
```

<span data-ttu-id="92060-273">Para ver um exemplo que usa essa t?cnica, confira [Inserir gr?ficos do Excel usando o Microsoft Graph em um Suplemento do PowerPoint](https://github.com/OfficeDev/PowerPoint-Add-in-Microsoft-Graph-ASPNET-InsertChart).</span><span class="sxs-lookup"><span data-stu-id="92060-273">For a sample that uses this technique, see [Insert Excel charts using Microsoft Graph in a PowerPoint Add-in](https://github.com/OfficeDev/PowerPoint-Add-in-Microsoft-Graph-ASPNET-InsertChart).</span></span>

<span data-ttu-id="92060-274">O c?digo na janela de di?logo pode analisar a URL e ler o valor do par?metro.</span><span class="sxs-lookup"><span data-stu-id="92060-274">Code in your dialog window can parse the URL and read the parameter value.</span></span>

> [!NOTE]
> <span data-ttu-id="92060-p139">O Office adiciona automaticamente um par?metro de consulta chamado `_host_info` ? URL que ? transmitida para `displayDialogAsync`. Ele ? anexado ap?s os par?metros de consulta personalizados, se houver algum. Ele n?o ? anexado ?s URLs subsequentes para as quais a caixa de di?logo navega. No futuro, a Microsoft poder? alterar o conte?do desse valor ou remov?-lo completamente para que seu c?digo n?o consiga l?-lo. O mesmo valor ? adicionado ao armazenamento de sess?o da caixa de di?logo. Novamente, *seu c?digo n?o deve ler nem gravar esse valor*.</span><span class="sxs-lookup"><span data-stu-id="92060-p139">Office automatically adds a query parameter called `_host_info` to the URL that is passed to `displayDialogAsync`. (It is appended after your custom query parameters, if any. It is not appended to any subsequent URLs that the dialog box navigates to.) Microsoft may change the content of this value, or remove it entirely, in the future, so your code should not read it. The same value is added to the dialog box's session storage. Again, *your code should neither read nor write to this value*.</span></span>

## <a name="use-the-dialog-apis-to-show-a-video"></a><span data-ttu-id="92060-280">Use APIs de Caixa de Di?logo para exibir um v?deo</span><span class="sxs-lookup"><span data-stu-id="92060-280">Use the Dialog APIs to show a video</span></span>

<span data-ttu-id="92060-281">Para mostrar um v?deo em uma caixa de di?logo:</span><span class="sxs-lookup"><span data-stu-id="92060-281">To show a video in a dialog box:</span></span>

1.  <span data-ttu-id="92060-p140">Crie uma p?gina cujo ?nico conte?do seja um iframe. O atributo `src` dos pontos do iframe para um v?deo online. O protocolo da URL do v?deo deve ser HTTP**S**. Neste artigo, chamaremos esta p?gina de "video.dialogbox.html". Veja a seguir um exemplo da marca??o:</span><span class="sxs-lookup"><span data-stu-id="92060-p140">Create a page whose only content is an iframe. The `src` attribute of the iframe points to an online video. The protocol of the video's URL must be HTTP**S**. In this article we'll call this page "video.dialogbox.html". The following is an example of the markup:</span></span>

    ```HTML
    <iframe class="ms-firstrun-video__player"  width="640" height="360"
        src="https://www.youtube.com/embed/XVfOe5mFbAE?rel=0&autoplay=1"
        frameborder="0" allowfullscreen>
    </iframe>
    ```

2.  <span data-ttu-id="92060-287">A p?gina video.dialogbox.html deve estar no mesmo dom?nio que a p?gina de host.</span><span class="sxs-lookup"><span data-stu-id="92060-287">The video.dialogbox.html page must be in the same domain as the host page.</span></span>
3.  <span data-ttu-id="92060-288">Use uma chamada de `displayDialogAsync` na p?gina host para abrir video.dialogbox.html.</span><span class="sxs-lookup"><span data-stu-id="92060-288">Use a call of `displayDialogAsync` in the host page to open video.dialogbox.html.</span></span>
4.  <span data-ttu-id="92060-p141">Se o suplemento precisar saber quando o usu?rio fecha a caixa de di?logo, registre um manipulador para o evento `DialogEventReceived` e manipule o evento 12006. Para mais detalhes, confira a se??o [Erros e eventos na janela de di?logo](#errors-and-events-in-the-dialog-window).</span><span class="sxs-lookup"><span data-stu-id="92060-p141">If your add-in needs to know when the user closes the dialog box, register a handler for the `DialogEventReceived` event and handle the 12006 event. For details, see the section [Errors and events in the dialog window](#errors-and-events-in-the-dialog-window).</span></span>

<span data-ttu-id="92060-291">Para ver um exemplo que usa um v?deo na caixa de di?logo, confira o [padr?o de design do video placemat](https://github.com/OfficeDev/Office-Add-in-UX-Design-Patterns-Code/tree/master/templates/first-run/video-placemat) no reposit?rio [padr?es de design da experi?ncia do usu?rio para Suplementos do Office](https://github.com/OfficeDev/Office-Add-in-UX-Design-Patterns-Code).</span><span class="sxs-lookup"><span data-stu-id="92060-291">For a sample that shows a video in a dialog box, see the [video placemat design pattern](https://github.com/OfficeDev/Office-Add-in-UX-Design-Patterns-Code/tree/master/templates/first-run/video-placemat) in the [UX design patterns for Office Add-ins](https://github.com/OfficeDev/Office-Add-in-UX-Design-Patterns-Code) repo.</span></span>

![Captura de tela de um v?deo mostrando uma caixa de di?logo de suplemento](../images/video-placemats-dialog-open.png)

## <a name="use-the-dialog-apis-in-an-authentication-flow"></a><span data-ttu-id="92060-293">Use as APIs de Caixa de Di?logo em um fluxo de autentica??o</span><span class="sxs-lookup"><span data-stu-id="92060-293">Use the Dialog APIs in an authentication flow</span></span>

<span data-ttu-id="92060-294">O principal cen?rio das APIs de Caixa de di?logo ? habilitar a autentica??o com um provedor de identidade ou recurso que n?o permite que a p?gina de entrada abra em um Iframe, como uma Conta da Microsoft, o Office 365, o Google e o Facebook.</span><span class="sxs-lookup"><span data-stu-id="92060-294">A primary scenario for the Dialog APIs is to enable authentication with a resource or identity provider that does not allow its sign-in page to open in an Iframe, such as Microsoft Account, Office 365, Google, and Facebook.</span></span>

> [!NOTE]
> <span data-ttu-id="92060-p142">Ao usar as APIs de Di?logo para esse cen?rio, *n?o* use a op??o `displayInIframe: true` na chamada para `displayDialogAsync`. Confira [Tirar proveito de uma op??o de desempenho no Office Online](#take-advantage-of-a-performance-option-in-office-online) para obter detalhes sobre essa op??o anteriormente neste artigo.</span><span class="sxs-lookup"><span data-stu-id="92060-p142">When you are using the Dialog APIs for this scenario, do *not* use the `displayInIframe: true` option in the call to `displayDialogAsync`. See [Take advantage of a performance option in Office Online](#take-advantage-of-a-performance-option-in-office-online) previously in this article for details about this option.</span></span>

<span data-ttu-id="92060-297">O que vem a seguir ? um fluxo de autentica??o simples e t?pico:</span><span class="sxs-lookup"><span data-stu-id="92060-297">The following is a simple and typical authentication flow:</span></span>

1. <span data-ttu-id="92060-p143">A primeira p?gina que ? aberta na caixa de di?logo ? uma p?gina local (ou outro recurso) que est? hospedada no dom?nio do suplemento; ou seja, o dom?nio da janela do host. Essa p?gina pode ter uma ?nica interface de usu?rio que informa "Aguarde. Estamos redirecionando voc? para a p?gina onde poder? entrar no *NOME-DO-PROVEDOR*." O c?digo nessa p?gina constr?i a URL da p?gina de entrada do provedor de identidade usando as informa??es que s?o transmitidas para a caixa de di?logo, conforme descrito em [Transmitir informa??es para a caixa de di?logo](#pass-information-to-the-dialog-box).</span><span class="sxs-lookup"><span data-stu-id="92060-p143">The first page that opens in the dialog box is a local page (or other resource) that is hosted in the add-in's domain; that is, the host window's domain. This page can have a simple UI that says "Please wait, we are redirecting you to the page where you can sign in to *NAME-OF-PROVIDER*." Code in this page constructs the URL of the identity provider's sign-in page by using information that is passed to the dialog box as described in [Pass information to the dialog box](#pass-information-to-the-dialog-box).</span></span>
2. <span data-ttu-id="92060-p144">A janela de di?logo redireciona ent?o para a p?gina de entrada. A URL inclui um par?metro de consulta que informa o provedor de identidade para redirecionar a janela de di?logo depois que o usu?rio entrar em uma p?gina espec?fica. Neste artigo, chamaremos essa p?gina de "redirectPage.html". (*Essa p?gina deve estar no mesmo dom?nio que a janela do host*, j? que a ?nica maneira de a janela de di?logo transmitir os resultados da tentativa de entrada ? usar uma chamada de `messageParent`, que s? pode ser chamada em uma p?gina com o mesmo dom?nio da janela do host.)</span><span class="sxs-lookup"><span data-stu-id="92060-p144">The dialog window then redirects to the sign-in page. The URL includes a query parameter that tells the identity provider to redirect the dialog window, after the user signs in, to a specific page. In this article, we'll call this page "redirectPage.html". (*This must be a page in the same domain as the host window*, because the only way for the dialog window to pass the results of the sign-in attempt is with a call of `messageParent`, which can only be called on a page with the same domain as the host window.)</span></span>
2. <span data-ttu-id="92060-p145">O servi?o do provedor de identidade processa a solicita??o GET recebida na janela de di?logo. Se o usu?rio j? estiver conectado, ele imediatamente redirecionar? a janela para redirectPage.html e incluir? os dados do usu?rio como um par?metro de consulta. Se o usu?rio ainda n?o tiver entrado, a p?gina de entrada do provedor aparecer? na janela para que o usu?rio possa entrar. Para a maioria dos provedores, se o usu?rio n?o consegue entrar com ?xito, o provedor mostra uma p?gina de erro na janela de di?logo e n?o redireciona para redirectPage.html. O usu?rio precisa fechar a janela selecionando o **X** no canto. Se o usu?rio entrar com ?xito, a janela de di?logo ser? redirecionada para redirectPage.html e os dados do usu?rio ser?o inclu?dos como um par?metro de consulta.</span><span class="sxs-lookup"><span data-stu-id="92060-p145">The identity provider's service processes the incoming GET request from the dialog window. If the user is already logged on, it immediately redirects the window to redirectPage.html and includes user data as a query parameter. If the user is not already signed in, the provider's sign-in page appears in the window, and the user signs in. For most providers, if the user cannot sign in successfully, the provider shows an error page in the dialog window and does not redirect to redirectPage.html. The user must close the window by selecting the **X** in the corner. If the user successfully signs in, the dialog window is redirected to redirectPage.html and user data is included as a query parameter.</span></span>
3. <span data-ttu-id="92060-311">Quando a p?gina redirectPage.html ? aberta, ela chama `messageParent` para relatar o ?xito ou falha na p?gina host e opcionalmente tamb?m informar dados do usu?rio ou dados de erro.</span><span class="sxs-lookup"><span data-stu-id="92060-311">When the redirectPage.html page opens, it calls `messageParent` to report the success or failure to the host page and optionally also report user data or error data.</span></span>
4. <span data-ttu-id="92060-312">O evento `DialogMessageReceived` ? acionado na p?gina host e seu manipulador fecha a janela de di?logo e, opcionalmente, faz outro processamento da mensagem.</span><span class="sxs-lookup"><span data-stu-id="92060-312">The `DialogMessageReceived` event fires in the host page and its handler closes the dialog window and optionally does other processing of the message.</span></span>

<span data-ttu-id="92060-313">Para ver suplementos de exemplo que usam esse padr?o, confira:</span><span class="sxs-lookup"><span data-stu-id="92060-313">For sample add-ins that use this pattern, see:</span></span>

- <span data-ttu-id="92060-p146">[Inserir gr?ficos do Excel usando o Microsoft Graph em um Suplemento do PowerPoint](https://github.com/OfficeDev/PowerPoint-Add-in-Microsoft-Graph-ASPNET-InsertChart): O recurso que ? inicialmente aberto na janela de di?logo ? um m?todo controlador que n?o tem seu pr?prio modo de exibi??o. Ele redireciona para a p?gina de entrada do Office 365.</span><span class="sxs-lookup"><span data-stu-id="92060-p146">[Insert Excel charts using Microsoft Graph in a PowerPoint Add-in](https://github.com/OfficeDev/PowerPoint-Add-in-Microsoft-Graph-ASPNET-InsertChart): The resource that is initially opened in the dialog window is a controller method that has no view of its own. It redirects to the Office 365 sign in page.</span></span>
- <span data-ttu-id="92060-316">[Autentica??o de Cliente do Office 365 de Suplementos do Office para AngularJS](https://github.com/OfficeDev/Word-Add-in-AngularJS-Client-OAuth): O recurso que ? inicialmente aberto na janela de di?logo ? uma p?gina.</span><span class="sxs-lookup"><span data-stu-id="92060-316">[Office Add-in Office 365 Client Authentication for AngularJS](https://github.com/OfficeDev/Word-Add-in-AngularJS-Client-OAuth): The resource that is initially opened in the dialog window is a page.</span></span>

#### <a name="support-multiple-identity-providers"></a><span data-ttu-id="92060-317">Prestar suporte a v?rios provedores de identidade</span><span class="sxs-lookup"><span data-stu-id="92060-317">Support multiple identity providers</span></span>

<span data-ttu-id="92060-p147">Se seu suplemento oferece ao usu?rio uma variedade de op??es de provedores, como a Conta da Microsoft, o Google ou o Facebook, voc? precisa de uma primeira p?gina local (confira a se??o anterior) que forne?a uma interface de usu?rio para a escolha de um provedor. A escolha do provedor acionar? a constru??o da URL de entrada e seu redirecionamento.</span><span class="sxs-lookup"><span data-stu-id="92060-p147">If your add-in gives the user a choice of providers, such as Microsoft Account, Google, or Facebook, you need a local first page (see preceding section) that provides a UI for the user to select a provider. Selection triggers the construction of the sign-in URL and redirection to it.</span></span>

<span data-ttu-id="92060-320">Para ver um exemplo que usa esse padr?o, confira [Suplemento do Office que usa o Servi?o Auth0 para Facilitar o Login Social](https://github.com/OfficeDev/Office-Add-in-Auth0).</span><span class="sxs-lookup"><span data-stu-id="92060-320">For a sample that uses this pattern, see [Office Add-in that uses the Auth0 Service to Simplify Social Login](https://github.com/OfficeDev/Office-Add-in-Auth0).</span></span>

#### <a name="authorization-of-the-add-in-to-an-external-resource"></a><span data-ttu-id="92060-321">Autoriza??o do suplemento para um recurso externo</span><span class="sxs-lookup"><span data-stu-id="92060-321">Authorization of the add-in to an external resource</span></span>

<span data-ttu-id="92060-p148">Na Web moderna, os aplicativos Web s?o entidades de seguran?a, como os usu?rios, e o aplicativo tem sua pr?pria identidade e permiss?es para recursos online, como o Office 365, Google Plus, Facebook ou LinkedIn. O aplicativo ? registrado no provedor de recursos antes da implanta??o. O registro inclui:</span><span class="sxs-lookup"><span data-stu-id="92060-p148">In the modern web, web applications are security principals just as users are, and the application has its own identity and permissions to an online resource such as Office 365, Google Plus, Facebook, or LinkedIn. The application is registered with the resource provider before it is deployed. The registration includes:</span></span>

- <span data-ttu-id="92060-325">Uma lista das permiss?es que o aplicativo precisa para usar recursos de um usu?rio.</span><span class="sxs-lookup"><span data-stu-id="92060-325">A list of the permissions that the application needs to a user's resources.</span></span>
- <span data-ttu-id="92060-326">Uma URL para a qual o servi?o do recurso deve retornar um token de acesso quando o aplicativo acessa o servi?o.</span><span class="sxs-lookup"><span data-stu-id="92060-326">A URL to which the resource service should return an access token when the application accesses the service.</span></span>  

<span data-ttu-id="92060-p149">Quando um usu?rio invoca uma fun??o no aplicativo que acessa os dados do usu?rio no servi?o do recurso, ele ? solicitado a entrar no servi?o e a conceder ao aplicativo as permiss?es necess?rias para os recursos do usu?rio. Em seguida, o servi?o redireciona a janela de entrada para a URL previamente registrada e transmite o token de acesso. O aplicativo usa o token de acesso para acessar os recursos do usu?rio.</span><span class="sxs-lookup"><span data-stu-id="92060-p149">When a user invokes a function in the application that accesses the user's data in the resource service, they are prompted to sign in to the service and then prompted to grant the application the permissions it needs to the user's resources. The service then redirects the sign-in window to the previously registered URL and passes the access token. The application uses the access token to access the user's resources.</span></span>

<span data-ttu-id="92060-p150">Voc? pode usar as APIs de Caixa de Di?logo para gerenciar esse processo usando um fluxo semelhante ?quele descrito para os usu?rios entrarem. As ?nicas diferen?as s?o:</span><span class="sxs-lookup"><span data-stu-id="92060-p150">You can use the Dialog APIs to manage this process by using a flow that is similar to the one described for users to sign in. The only differences are:</span></span>

- <span data-ttu-id="92060-332">Se o usu?rio ainda n?o tiver concedido ao aplicativo as permiss?es necess?rias, ele ser? solicitada a faz?-lo na caixa de di?logo ap?s entrar.</span><span class="sxs-lookup"><span data-stu-id="92060-332">If the user hasn't previously granted the application the permissions it needs, she is prompted to do so in the dialog box after signing in.</span></span>
- <span data-ttu-id="92060-p151">A janela de di?logo envia o token de acesso ? janela do host usando `messageParent` para enviar o token de acesso em formato de cadeia de caracteres ou armazenando o token de acesso em um local onde a janela do host poder? recuper?-lo. O token tem um limite de tempo, mas enquanto durar, a janela do host poder us?-lo para acessar recursos do usu?rio de forma direta, sem outras solicita??es.</span><span class="sxs-lookup"><span data-stu-id="92060-p151">The dialog window sends the access token to the host window either by using `messageParent` to send the stringified access token or by storing the access token where the host window can retrieve it. The token has a time limit, but while it lasts, the host window can use it to directly access the user's resources without any further prompting.</span></span>

<span data-ttu-id="92060-335">Os exemplos a seguir usam as APIs de Caixa de di?logo para essa finalidade:</span><span class="sxs-lookup"><span data-stu-id="92060-335">The following samples use the Dialog APIs for this purpose:</span></span>
- <span data-ttu-id="92060-336">[Inserir gr?ficos do Excel usando o Microsoft Graph em um Suplemento do PowerPoint](https://github.com/OfficeDev/PowerPoint-Add-in-Microsoft-Graph-ASPNET-InsertChart): armazena o token de acesso em um banco de dados.</span><span class="sxs-lookup"><span data-stu-id="92060-336">[Insert Excel charts using Microsoft Graph in a PowerPoint Add-in](https://github.com/OfficeDev/PowerPoint-Add-in-Microsoft-Graph-ASPNET-InsertChart) - Stores the access token in a database.</span></span>
- [<span data-ttu-id="92060-337">Suplemento do Office que usa o Servi?o do OAuth.io para Simplificar o Acesso a Servi?os Populares Online</span><span class="sxs-lookup"><span data-stu-id="92060-337">Office Add-in that uses the OAuth.io Service to Simplify Access to Popular Online Services</span></span>](https://github.com/OfficeDev/Office-Add-in-OAuth.io)

<span data-ttu-id="92060-338">Para mais informa??es sobre a autentica??o e autoriza??o em suplementos, consulte:</span><span class="sxs-lookup"><span data-stu-id="92060-338">For more information about authentication and authorization in add-ins, see:</span></span>
- [<span data-ttu-id="92060-339">Autorizar servi?os externos no seu Suplemento do Office</span><span class="sxs-lookup"><span data-stu-id="92060-339">Authorize external services in your Office Add-in</span></span>](auth-external-add-ins.md)
- [<span data-ttu-id="92060-340">Biblioteca de Auxiliares da API JavaScript para Office</span><span class="sxs-lookup"><span data-stu-id="92060-340">Office JavaScript API Helpers library</span></span>](https://github.com/OfficeDev/office-js-helpers)


## <a name="use-the-office-dialog-api-with-single-page-applications-and-client-side-routing"></a><span data-ttu-id="92060-341">Usar a API de Caixa de di?logo para Office com aplicativos de p?gina ?nica e roteamento do lado do cliente</span><span class="sxs-lookup"><span data-stu-id="92060-341">Use the Office Dialog API with single-page applications and client-side routing</span></span>

<span data-ttu-id="92060-342">Se seu suplemento usa o roteamento do lado do cliente, como os aplicativos de p?gina ?nica geralmente fazem, voc? tem a op??o de transmitir a URL de um roteamento para o m?todo [displayDialogAsync](http://dev.office.com/reference/add-ins/shared/officeui.displaydialogasync), em vez da URL de uma p?gina HTML completa e separada.</span><span class="sxs-lookup"><span data-stu-id="92060-342">If your add-in uses client-side routing, as single-page applications typically do, you have the option to pass the URL of a route to the [displayDialogAsync](http://dev.office.com/reference/add-ins/shared/officeui.displaydialogasync) method, instead of the URL of a complete and separate HTML page.</span></span>

> [!IMPORTANT]
><span data-ttu-id="92060-p152">A caixa de di?logo est? em uma nova janela com seu pr?prio contexto de execu??o. Se voc? transmitir uma rota, sua p?gina de base e todos os c?digos de inicializa??o e bootstrapping ser?o executados novamente nesse novo contexto e todas as vari?veis ser?o definidas com seus valores iniciais na caixa de di?logo. Portanto, essa t?cnica inicia uma segunda inst?ncia do aplicativo na janela de di?logo. O c?digo que altera as vari?veis na janela de di?logo n?o altera a vers?o do painel tarefas das mesmas vari?veis. De forma semelhante, a janela de di?logo tem seu pr?prio armazenamento de sess?o que n?o pode ser acessado a partir do c?digo no painel de tarefas.</span><span class="sxs-lookup"><span data-stu-id="92060-p152">The dialog box is in a new window with its own execution context. If you pass a route, your base page and all its initialization and bootstrapping code run again in this new context, and any variables are set to their initial values in the dialog window. So this technique launches a second instance of your application in the dialog window. Code that changes variables in the dialog window does not change the task pane version of the same variables. Similarly, the dialog window has its own session storage, which is not accessible from code in the task pane.</span></span>
