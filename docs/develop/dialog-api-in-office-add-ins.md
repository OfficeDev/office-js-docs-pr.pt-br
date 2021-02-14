---
title: Usar a API da Caixa de Diálogo do Office nos suplementos do Office
description: Conheça as noções básicas da criação de uma caixa de diálogo em um Complemento do Office.
ms.date: 01/28/2021
localization_priority: Normal
ms.openlocfilehash: 9061b4c048a133572e615152d61df611e5f15068
ms.sourcegitcommit: ccc0a86d099ab4f5ef3d482e4ae447c3f9b818a3
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 02/14/2021
ms.locfileid: "50237858"
---
# <a name="use-the-office-dialog-api-in-office-add-ins"></a><span data-ttu-id="00ecb-103">Usar a API de diálogo do Office em suplementos do Office</span><span class="sxs-lookup"><span data-stu-id="00ecb-103">Use the Office dialog API in Office Add-ins</span></span>

<span data-ttu-id="00ecb-104">Você pode usar a [API de Caixa de diálogo do Office](/javascript/api/office/office.ui) para abrir caixas de diálogo no seu Suplemento do Office.</span><span class="sxs-lookup"><span data-stu-id="00ecb-104">You can use the [Office dialog API](/javascript/api/office/office.ui) to open dialog boxes in your Office Add-in.</span></span> <span data-ttu-id="00ecb-105">Este artigo fornece orientações para usar a API de Caixa de diálogo em seu Suplemento do Office.</span><span class="sxs-lookup"><span data-stu-id="00ecb-105">This article provides guidance for using the dialog API in your Office Add-in.</span></span>

> [!NOTE]
> <span data-ttu-id="00ecb-106">Para informações sobre os programas para os quais a API de Caixa de Diálogo tem suporte no momento, confira [Conjuntos de requisitos da API de Caixa de Diálogo](../reference/requirement-sets/dialog-api-requirement-sets.md).</span><span class="sxs-lookup"><span data-stu-id="00ecb-106">For information about where the Dialog API is currently supported, see [Dialog API requirement sets](../reference/requirement-sets/dialog-api-requirement-sets.md).</span></span> <span data-ttu-id="00ecb-107">Atualmente, a API da Caixa de Diálogo tem suporte para Excel, PowerPoint e Word.</span><span class="sxs-lookup"><span data-stu-id="00ecb-107">The Dialog API is currently supported for Excel, PowerPoint, and Word.</span></span> <span data-ttu-id="00ecb-108">O suporte do Outlook está incluído em vários conjuntos de requisitos de Caixa de &mdash; Correio, consulte a referência de API para obter mais detalhes.</span><span class="sxs-lookup"><span data-stu-id="00ecb-108">Outlook support is included across various Mailbox requirement sets&mdash;see the API reference for more details.</span></span>

<span data-ttu-id="00ecb-109">Um cenário fundamental para a API de Caixa de Diálogo é habilitar a autenticação com um recurso como o Google, o Facebook ou o Microsoft Graph.</span><span class="sxs-lookup"><span data-stu-id="00ecb-109">A primary scenario for the Dialog API is to enable authentication with a resource such as Google, Facebook, or Microsoft Graph.</span></span> <span data-ttu-id="00ecb-110">Para saber mais, confira [ autenticação com APIs de Caixa de Diálogo do Office](auth-with-office-dialog-api.md) *depois* que você se familiarizar com este artigo.</span><span class="sxs-lookup"><span data-stu-id="00ecb-110">For more information, see [Authenticate with the Office dialog API](auth-with-office-dialog-api.md) *after* you are familiar with this article.</span></span>

<span data-ttu-id="00ecb-111">Considere abrir uma caixa de diálogo em um painel de tarefas, suplemento de conteúdo ou [comando de suplemento](../design/add-in-commands.md) para fazer o seguinte:</span><span class="sxs-lookup"><span data-stu-id="00ecb-111">Consider opening a dialog box from a task pane or content add-in or [add-in command](../design/add-in-commands.md) to do the following:</span></span>

- <span data-ttu-id="00ecb-112">Exibir páginas de entrada que não podem ser abertas diretamente em um painel de tarefas.</span><span class="sxs-lookup"><span data-stu-id="00ecb-112">Display sign in pages that cannot be opened directly in a task pane.</span></span>
- <span data-ttu-id="00ecb-113">Fornecer mais espaço na tela, ou até uma tela inteira, para algumas tarefas no seu suplemento.</span><span class="sxs-lookup"><span data-stu-id="00ecb-113">Provide more screen space, or even a full screen, for some tasks in your add-in.</span></span>
- <span data-ttu-id="00ecb-114">Hospedar um vídeo que seria muito pequeno se fosse confinado em um painel de tarefas.</span><span class="sxs-lookup"><span data-stu-id="00ecb-114">Host a video that would be too small if confined to a task pane.</span></span>

> [!NOTE]
> <span data-ttu-id="00ecb-115">Como a sobreposição de elementos de IU não são recomendáveis, evite abrir uma caixa de diálogo em um painel de tarefas a menos que seu cenário o obrigue a fazer isso.</span><span class="sxs-lookup"><span data-stu-id="00ecb-115">Because overlapping UI elements are discouraged, avoid opening a dialog box from a task pane unless your scenario requires it.</span></span> <span data-ttu-id="00ecb-116">Ao considerar como usar a área de superfície de um painel de tarefas, observe que painéis de tarefas podem ter guias.</span><span class="sxs-lookup"><span data-stu-id="00ecb-116">When you consider how to use the surface area of a task pane, note that task panes can be tabbed.</span></span> <span data-ttu-id="00ecb-117">Para ver um exemplo de um painel de tarefas com guias, confira o exemplo [do JavaScript SalesTracker para o](https://github.com/OfficeDev/Excel-Add-in-JavaScript-SalesTracker) Complemento do Excel.</span><span class="sxs-lookup"><span data-stu-id="00ecb-117">For an example of a tabbed task pane, see the [Excel Add-in JavaScript SalesTracker](https://github.com/OfficeDev/Excel-Add-in-JavaScript-SalesTracker) sample.</span></span>

<span data-ttu-id="00ecb-118">A imagem abaixo mostra um exemplo de uma caixa de diálogo.</span><span class="sxs-lookup"><span data-stu-id="00ecb-118">The following image shows an example of a dialog box.</span></span>

![Captura de tela mostrando a caixa de diálogo com três opções de entrada exibidas na frente do Word](../images/auth-o-dialog-open.png)

<span data-ttu-id="00ecb-120">A caixa de diálogo sempre abre no centro da tela.</span><span class="sxs-lookup"><span data-stu-id="00ecb-120">Note that the dialog box always opens in the center of the screen.</span></span> <span data-ttu-id="00ecb-121">O usuário pode movê-la e redimensioná-la.</span><span class="sxs-lookup"><span data-stu-id="00ecb-121">The user can move and resize it.</span></span> <span data-ttu-id="00ecb-122">A janela *não émodal*– um usuário pode continuar a interagir com o documento no aplicativo do Office e com a página no painel de tarefas, se houver uma.</span><span class="sxs-lookup"><span data-stu-id="00ecb-122">The window is *nonmodal*--a user can continue to interact with both the document in the Office application and with the page in the task pane, if there is one.</span></span>

## <a name="open-a-dialog-box-from-a-host-page"></a><span data-ttu-id="00ecb-123">Abrir uma caixa de diálogo em uma página de host</span><span class="sxs-lookup"><span data-stu-id="00ecb-123">Open a dialog box from a host page</span></span>

<span data-ttu-id="00ecb-124">As APIs JavaScript para Office incluem um objeto[Dialog](/javascript/api/office/office.dialog) e duas funções no [namespace Office.context.ui](/javascript/api/office/office.ui).</span><span class="sxs-lookup"><span data-stu-id="00ecb-124">The Office JavaScript APIs include a [Dialog](/javascript/api/office/office.dialog) object and two functions in the [Office.context.ui namespace](/javascript/api/office/office.ui).</span></span>

<span data-ttu-id="00ecb-125">Para abrir uma caixa de diálogo, seu código, geralmente uma página no painel de tarefas chama o método [displayDialogAsync](/javascript/api/office/office.ui) e transmite a ele a URL do recurso que você deseja abrir.</span><span class="sxs-lookup"><span data-stu-id="00ecb-125">To open a dialog box, your code, typically a page in a task pane, calls the [displayDialogAsync](/javascript/api/office/office.ui) method and passes to it the URL of the resource that you want to open.</span></span> <span data-ttu-id="00ecb-126">A página em que esse método é chamado é conhecida como "página host".</span><span class="sxs-lookup"><span data-stu-id="00ecb-126">The page on which this method is called is known as the "host page".</span></span> <span data-ttu-id="00ecb-127">Por exemplo, se você chamar esse método no script index.html em um painel de tarefas, index.html será a página do host da caixa de diálogo que o método abre.</span><span class="sxs-lookup"><span data-stu-id="00ecb-127">For example, if you call this method in script on index.html in a task pane, then index.html is the host page of the dialog box that the method opens.</span></span>

<span data-ttu-id="00ecb-128">O recurso aberto na página de diálogo geralmente é uma página, mas pode ser um método controlador em um aplicativo MVC, uma rota, um método de serviço Web ou qualquer outro recurso.</span><span class="sxs-lookup"><span data-stu-id="00ecb-128">The resource that is opened in the dialog box is usually a page, but it can be a controller method in an MVC application, a route, a web service method, or any other resource.</span></span> <span data-ttu-id="00ecb-129">Neste artigo, 'página' ou 'site' refere-se ao recurso na caixa de diálogo.</span><span class="sxs-lookup"><span data-stu-id="00ecb-129">In this article, 'page' or 'website' refers to the resource in the dialog box.</span></span> <span data-ttu-id="00ecb-130">O código a seguir é um exemplo simples:</span><span class="sxs-lookup"><span data-stu-id="00ecb-130">The following code is a simple example:</span></span>

```js
Office.context.ui.displayDialogAsync('https://myAddinDomain/myDialog.html');
```

> [!NOTE]
> - <span data-ttu-id="00ecb-131">A URL usa o protocolo HTTP **S**.</span><span class="sxs-lookup"><span data-stu-id="00ecb-131">The URL uses the HTTP **S** protocol.</span></span> <span data-ttu-id="00ecb-132">Isso é obrigatório para todas as páginas carregadas em uma caixa diálogo, não apenas para a primeira página carregada.</span><span class="sxs-lookup"><span data-stu-id="00ecb-132">This is mandatory for all pages loaded in a dialog box, not just the first page loaded.</span></span>
> - <span data-ttu-id="00ecb-133">A caixa de diálogo é igual ao domínio da página de host, que pode ser a página em um painel de tarefas ou o [arquivo de função](../reference/manifest/functionfile.md) de um comando de suplemento.</span><span class="sxs-lookup"><span data-stu-id="00ecb-133">The dialog box's domain is the same as the domain of the host page, which can be the page in a task pane or the [function file](../reference/manifest/functionfile.md) of an add-in command.</span></span> <span data-ttu-id="00ecb-134">Isso é necessário: a página, o método do controlador ou outro recurso que é passado para o método `displayDialogAsync` deve estar no mesmo domínio que a página de host.</span><span class="sxs-lookup"><span data-stu-id="00ecb-134">This is required: the page, controller method, or other resource that is passed to the `displayDialogAsync` method must be in the same domain as the host page.</span></span>

> [!IMPORTANT]
> <span data-ttu-id="00ecb-135">A página de host e o recurso que abrem na caixa de diálogo devem ter o mesmo domínio inteiro.</span><span class="sxs-lookup"><span data-stu-id="00ecb-135">The host page and the resource that opens in the dialog box must have the same full domain.</span></span> <span data-ttu-id="00ecb-136">Se você tentar passar `displayDialogAsync` para um subdomínio do domínio do suplemento, ele não funcionará.</span><span class="sxs-lookup"><span data-stu-id="00ecb-136">If you attempt to pass `displayDialogAsync` a subdomain of the add-in's domain, it will not work.</span></span> <span data-ttu-id="00ecb-137">O domínio completo, incluindo qualquer subdomínio, deve corresponder.</span><span class="sxs-lookup"><span data-stu-id="00ecb-137">The full domain, including any subdomain, must match.</span></span>

<span data-ttu-id="00ecb-138">Após o carregamento da primeira página (ou de outro recurso), um usuário pode usar links ou outra interface de usuário para qualquer site (ou outro recurso) que usa HTTPS.</span><span class="sxs-lookup"><span data-stu-id="00ecb-138">After the first page (or other resource) is loaded, a user can use links or other UI to navigate to any website (or other resource) that uses HTTPS.</span></span> <span data-ttu-id="00ecb-139">Também é possível criar a primeira página para redirecionar imediatamente para outro site.</span><span class="sxs-lookup"><span data-stu-id="00ecb-139">You can also design the first page to immediately redirect to another site.</span></span>

<span data-ttu-id="00ecb-140">Por padrão, a caixa de diálogo ocupará 80% da altura e da largura na tela do dispositivo, mas você pode definir porcentagens diferentes. Basta transmitir um objeto de configuração para o método, como mostra o exemplo a seguir.</span><span class="sxs-lookup"><span data-stu-id="00ecb-140">By default, the dialog box will occupy 80% of the height and width of the device screen, but you can set different percentages by passing a configuration object to the method, as shown in the following example:</span></span>

```js
Office.context.ui.displayDialogAsync('https://myDomain/myDialog.html', {height: 30, width: 20});
```

<span data-ttu-id="00ecb-141">Para ver um suplemento de exemplo que faz isso, confira [Exemplo de API de Caixa de diálogo do Suplemento do Office](https://github.com/OfficeDev/Office-Add-in-Dialog-API-Simple-Example).</span><span class="sxs-lookup"><span data-stu-id="00ecb-141">For a sample add-in that does this, see [Office Add-in Dialog API Example](https://github.com/OfficeDev/Office-Add-in-Dialog-API-Simple-Example).</span></span> <span data-ttu-id="00ecb-142">Para obter mais exemplos que `displayDialogAsync` usam, consulte [Exemplos](#samples).</span><span class="sxs-lookup"><span data-stu-id="00ecb-142">For more samples that use `displayDialogAsync`, see [Samples](#samples).</span></span>

<span data-ttu-id="00ecb-p113">Defina os dois valores como 100% para ter uma verdadeira experiência de tela inteira. O máximo real é 99,5%, e a janela ainda poderá ser movida e redimensionada.</span><span class="sxs-lookup"><span data-stu-id="00ecb-p113">Set both values to 100% to get what is effectively a full screen experience. (The effective maximum is 99.5%, and the window is still moveable and resizable.)</span></span>

> [!NOTE]
> <span data-ttu-id="00ecb-p114">Apenas uma caixa de diálogo pode ser aberta em uma janela do host. Tentar abrir outra caixa de diálogo gera um erro. Portanto, por exemplo, se um usuário abrir uma caixa de diálogo no painel de tarefas, ele não poderá abrir uma segunda caixa de diálogo em uma página diferente no painel de tarefas. No entanto, quando uma caixa de diálogo é aberta em um [comando de suplemento](../design/add-in-commands.md), o comando abre um arquivo HTML novo (mas não visto) sempre que ele é selecionado. Isso cria uma nova janela do host (não vista) para que cada janela possa iniciar sua própria caixa de diálogo. Para obter mais informações, confira [Erros de displayDialogAsync](dialog-handle-errors-events.md#errors-from-displaydialogasync).</span><span class="sxs-lookup"><span data-stu-id="00ecb-p114">You can open only one dialog box from a host window. An attempt to open another dialog box generates an error. For example, if a user opens a dialog box from a task pane, she cannot open a second dialog box, from a different page in the task pane. However, when a dialog box is opened from an [add-in command](../design/add-in-commands.md), the command opens a new (but unseen) HTML file each time it is selected. This creates a new (unseen) host window, so each such window can launch its own dialog box. For more information, see [Errors from displayDialogAsync](dialog-handle-errors-events.md#errors-from-displaydialogasync).</span></span>

### <a name="take-advantage-of-a-performance-option-in-office-on-the-web"></a><span data-ttu-id="00ecb-151">Aproveite uma opção de desempenho no Office na Web</span><span class="sxs-lookup"><span data-stu-id="00ecb-151">Take advantage of a performance option in Office on the web</span></span>

<span data-ttu-id="00ecb-152">A propriedade `displayInIframe` é uma propriedade adicional no objeto de configuração que você pode passar para o`displayDialogAsync`.</span><span class="sxs-lookup"><span data-stu-id="00ecb-152">The `displayInIframe` property is an additional property in the configuration object that you can pass to `displayDialogAsync`.</span></span> <span data-ttu-id="00ecb-153">Quando essa propriedade for definida como `true` e o suplemento estiver em execução em um documento aberto no Office Online, a caixa de diálogo será aberta como um iframe flutuante, em vez de uma janela independente, o que faz com que ela seja aberta mais rapidamente.</span><span class="sxs-lookup"><span data-stu-id="00ecb-153">When this property is set to `true`, and the add-in is running in a document opened in Office on the web, the dialog box will open as a floating iframe rather than an independent window, which makes it open faster.</span></span> <span data-ttu-id="00ecb-154">Este é um exemplo:</span><span class="sxs-lookup"><span data-stu-id="00ecb-154">The following is an example:</span></span>

```js
Office.context.ui.displayDialogAsync('https://myDomain/myDialog.html', {height: 30, width: 20, displayInIframe: true});
```

<span data-ttu-id="00ecb-155">O valor padrão é `false`, que é o mesmo que omitir a propriedade inteiramente.</span><span class="sxs-lookup"><span data-stu-id="00ecb-155">The default value is `false`, which is the same as omitting the property entirely.</span></span> <span data-ttu-id="00ecb-156">Se o suplemento não estiver sendo executado no Office Online, o `displayInIframe` será ignorado.</span><span class="sxs-lookup"><span data-stu-id="00ecb-156">If the add-in is not running in Office on the web, the `displayInIframe` is ignored.</span></span>

> [!NOTE]
> <span data-ttu-id="00ecb-157">Você **não** deverá usar `displayInIframe: true` se a caixa de diálogo redirecionar a qualquer ponto para uma página que não possa ser aberta em um iframe.</span><span class="sxs-lookup"><span data-stu-id="00ecb-157">You should **not** use `displayInIframe: true` if the dialog box will at any point redirect to a page that cannot be opened in an iframe.</span></span> <span data-ttu-id="00ecb-158">Por exemplo, as páginas de entrada de muitos serviços Web populares, como a conta do Google e da Microsoft, não podem ser abertas em um iframe.</span><span class="sxs-lookup"><span data-stu-id="00ecb-158">For example, the sign in pages of many popular web services, such as Google and Microsoft account, cannot be opened in an iframe.</span></span>

## <a name="send-information-from-the-dialog-box-to-the-host-page"></a><span data-ttu-id="00ecb-159">Envie informações da caixa de diálogo para a página host</span><span class="sxs-lookup"><span data-stu-id="00ecb-159">Send information from the dialog box to the host page</span></span>

<span data-ttu-id="00ecb-160">A caixa de diálogo não pode se comunicar com a página host no painel de tarefas, a menos que:</span><span class="sxs-lookup"><span data-stu-id="00ecb-160">The dialog box cannot communicate with the host page in the task pane unless:</span></span>

- <span data-ttu-id="00ecb-161">A página atual na caixa de diálogo esteja no mesmo domínio da página host.</span><span class="sxs-lookup"><span data-stu-id="00ecb-161">The current page in the dialog box is in the same domain as the host page.</span></span>
- <span data-ttu-id="00ecb-162">A biblioteca da API JavaScript do Office é carregada na página.</span><span class="sxs-lookup"><span data-stu-id="00ecb-162">The Office JavaScript API library is loaded in the page.</span></span> <span data-ttu-id="00ecb-163">(Como qualquer página que usa a biblioteca da API JavaScript do Office, o script para a página deve atribuir um método à propriedade, embora possa `Office.initialize` ser um método vazio.</span><span class="sxs-lookup"><span data-stu-id="00ecb-163">(Like any page that uses the Office JavaScript API library, script for the page must assign a method to the `Office.initialize` property, although it can be an empty method.</span></span> <span data-ttu-id="00ecb-164">Para obter detalhes, [confira Inicializar o seu Complemento do Office.)](initialize-add-in.md)</span><span class="sxs-lookup"><span data-stu-id="00ecb-164">For details, see [Initialize your Office Add-in](initialize-add-in.md).)</span></span>

<span data-ttu-id="00ecb-165">O código na caixa de diálogo use a função [messageParent](/javascript/api/office/office.ui#messageparent-message-) para enviar uma mensagem de cadeia de caracteres ou um valor booliano para a página host.</span><span class="sxs-lookup"><span data-stu-id="00ecb-165">Code in the dialog box uses the [messageParent](/javascript/api/office/office.ui#messageparent-message-) function to send either a Boolean value or a string message to the host page.</span></span> <span data-ttu-id="00ecb-166">A cadeia de caracteres pode ser uma palavra, uma frase, um blob XML, um JSON em formato de cadeia de caracteres ou qualquer outra coisa que possa ser serializada em uma cadeia de caracteres.</span><span class="sxs-lookup"><span data-stu-id="00ecb-166">The string can be a word, sentence, XML blob, stringified JSON, or anything else that can be serialized to a string.</span></span> <span data-ttu-id="00ecb-167">Este é um exemplo:</span><span class="sxs-lookup"><span data-stu-id="00ecb-167">The following is an example:</span></span>

```js
if (loginSuccess) {
    Office.context.ui.messageParent(true);
}
```

> [!IMPORTANT]
> - <span data-ttu-id="00ecb-168">A função `messageParent` só pode ser chamada em uma página com o mesmo domínio (incluindo o protocolo e a porta) da página host.</span><span class="sxs-lookup"><span data-stu-id="00ecb-168">The `messageParent` function can only be called on a page with the same domain (including protocol and port) as the host page.</span></span>
> - <span data-ttu-id="00ecb-169">A `messageParent` função é uma das *duas* APIs do Office JS que podem ser chamadas na caixa de diálogo.</span><span class="sxs-lookup"><span data-stu-id="00ecb-169">The `messageParent` function is one of *only* two Office JS APIs that can be called in the dialog box.</span></span>
> - <span data-ttu-id="00ecb-170">A outra API JS que pode ser chamada na caixa de diálogo é `Office.context.requirements.isSetSupported` .</span><span class="sxs-lookup"><span data-stu-id="00ecb-170">The other JS API that can be called in the dialog box is `Office.context.requirements.isSetSupported`.</span></span> <span data-ttu-id="00ecb-171">Para saber mais sobre isso, confira [Especificar aplicativos do Office e requisitos de API.](specify-office-hosts-and-api-requirements.md)</span><span class="sxs-lookup"><span data-stu-id="00ecb-171">For information about it, see [Specify Office applications and API requirements](specify-office-hosts-and-api-requirements.md).</span></span> <span data-ttu-id="00ecb-172">No entanto, na caixa de diálogo, essa API não é suportada na compra única do Outlook 2016 (ou seja, a versão MSI).</span><span class="sxs-lookup"><span data-stu-id="00ecb-172">However, in the dialog box, this API isn't supported in Outlook 2016 one-time purchase (that is, the MSI version).</span></span>

<span data-ttu-id="00ecb-173">No próximo exemplo, `googleProfile` é uma versão em formato de cadeia de caracteres do perfil do Google do usuário.</span><span class="sxs-lookup"><span data-stu-id="00ecb-173">In the next example, `googleProfile` is a stringified version of the user's Google profile.</span></span>

```js
if (loginSuccess) {
    Office.context.ui.messageParent(googleProfile);
}
```

<span data-ttu-id="00ecb-p121">A página host deve ser configurada para receber a mensagem. Você pode fazer isso adicionando um parâmetro de retorno de chamada à chamada original de `displayDialogAsync`. O retorno de chamada atribui um manipulador ao evento `DialogMessageReceived`. Apresentamos um exemplo a seguir:</span><span class="sxs-lookup"><span data-stu-id="00ecb-p121">The host page must be configured to receive the message. You do this by adding a callback parameter to the original call of `displayDialogAsync`. The callback assigns a handler to the `DialogMessageReceived` event. The following is an example:</span></span>

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
> - <span data-ttu-id="00ecb-178">O Office transmite um objeto [AsyncResult](/javascript/api/office/office.asyncresult) para o retorno de chamada.</span><span class="sxs-lookup"><span data-stu-id="00ecb-178">Office passes an [AsyncResult](/javascript/api/office/office.asyncresult) object to the callback.</span></span> <span data-ttu-id="00ecb-179">Ele representa o resultado de tentativas de abrir a caixa de diálogo, </span><span class="sxs-lookup"><span data-stu-id="00ecb-179">It represents the result of the attempt to open the dialog box.</span></span> <span data-ttu-id="00ecb-180">Ela não representa o resultado de eventos na caixa diálogo.</span><span class="sxs-lookup"><span data-stu-id="00ecb-180">It does not represent the outcome of any events in the dialog box.</span></span> <span data-ttu-id="00ecb-181">Para saber mais sobre essa distinção, confira [Manipular erros e eventos](dialog-handle-errors-events.md).</span><span class="sxs-lookup"><span data-stu-id="00ecb-181">For more on this distinction, see [Handle errors and events](dialog-handle-errors-events.md).</span></span>
> - <span data-ttu-id="00ecb-182">A propriedade `value` do `asyncResult` é definida como um objeto [Dialog](/javascript/api/office/office.dialog) que existe na página host, não no contexto da execução da caixa de diálogo.</span><span class="sxs-lookup"><span data-stu-id="00ecb-182">The `value` property of the `asyncResult` is set to a [Dialog](/javascript/api/office/office.dialog) object, which exists in the host page, not in the dialog box's execution context.</span></span>
> - <span data-ttu-id="00ecb-p123">O `processMessage` é a função que manipula o evento. Você pode dar a ele o nome que desejar.</span><span class="sxs-lookup"><span data-stu-id="00ecb-p123">The `processMessage` is the function that handles the event. You can give it any name you want.</span></span>
> - <span data-ttu-id="00ecb-185">A variável `dialog` é declarada em um escopo mais amplo do que o retorno de chamada porque ela também é referenciada em `processMessage`.</span><span class="sxs-lookup"><span data-stu-id="00ecb-185">The `dialog` variable is declared at a wider scope than the callback because it is also referenced in `processMessage`.</span></span>

<span data-ttu-id="00ecb-186">Veja a seguir um exemplo simples de um manipulador para o evento `DialogMessageReceived`:</span><span class="sxs-lookup"><span data-stu-id="00ecb-186">The following is a simple example of a handler for the `DialogMessageReceived` event:</span></span>

```js
function processMessage(arg) {
    var messageFromDialog = JSON.parse(arg.message);
    showUserName(messageFromDialog.name);
}
```

> [!NOTE]
> - <span data-ttu-id="00ecb-187">O Office transmite o objeto `arg` para o manipulador.</span><span class="sxs-lookup"><span data-stu-id="00ecb-187">Office passes the `arg` object to the handler.</span></span> <span data-ttu-id="00ecb-188">Sua propriedade `message` é o booliano ou a cadeia de caracteres enviada pela chamada de `messageParent` na caixa de diálogo.</span><span class="sxs-lookup"><span data-stu-id="00ecb-188">Its `message` property is the Boolean or string sent by the call of `messageParent` in the dialog box.</span></span> <span data-ttu-id="00ecb-189">Neste exemplo, é uma representação em cadeia de caracteres do perfil de um usuário de um serviço como a conta da Microsoft ou o Google, portanto, ela é desserlizada de volta para um objeto com `JSON.parse` .</span><span class="sxs-lookup"><span data-stu-id="00ecb-189">In this example, it is a stringified representation of a user's profile from a service such as Microsoft account or Google, so it is deserialized back to an object with `JSON.parse`.</span></span>
> - <span data-ttu-id="00ecb-p125">A implementação de `showUserName` não é mostrada. Ela pode exibir uma mensagem de boas-vindas personalizada no painel de tarefas.</span><span class="sxs-lookup"><span data-stu-id="00ecb-p125">The `showUserName` implementation is not shown. It might display a personalized welcome message on the task pane.</span></span>

<span data-ttu-id="00ecb-192">Quando a interação do usuário com a caixa de diálogo for concluída, seu manipulador de mensagem fechará a caixa de diálogo, conforme mostrado neste exemplo.</span><span class="sxs-lookup"><span data-stu-id="00ecb-192">When the user interaction with the dialog box is completed, your message handler should close the dialog box, as shown in this example.</span></span>

```js
function processMessage(arg) {
    dialog.close();
    // message processing code goes here;
}
```

> [!NOTE]
> - <span data-ttu-id="00ecb-193">O objeto `dialog` deve ser o mesmo que é retornado pela chamada de `displayDialogAsync`.</span><span class="sxs-lookup"><span data-stu-id="00ecb-193">The `dialog` object must be the same one that is returned by the call of `displayDialogAsync`.</span></span>
> - <span data-ttu-id="00ecb-194">A chamada de `dialog.close` informa ao Office para fechar a caixa de diálogo imediatamente.</span><span class="sxs-lookup"><span data-stu-id="00ecb-194">The call of `dialog.close` tells Office to immediately close the dialog box.</span></span>

<span data-ttu-id="00ecb-195">Para ver um suplemento de exemplo que usa essas técnicas, confira [Exemplo de API de Caixa de diálogo do Suplemento do Office](https://github.com/OfficeDev/Office-Add-in-Dialog-API-Simple-Example).</span><span class="sxs-lookup"><span data-stu-id="00ecb-195">For a sample add-in that uses these techniques, see [Office Add-in Dialog API Example](https://github.com/OfficeDev/Office-Add-in-Dialog-API-Simple-Example).</span></span>

<span data-ttu-id="00ecb-p126">Se o suplemento precisa abrir uma página diferente do painel de tarefas depois de receber a mensagem, é possível usar o método `window.location.replace` (ou `window.location.href`) como a última linha do manipulador. Apresentamos um exemplo a seguir:</span><span class="sxs-lookup"><span data-stu-id="00ecb-p126">If the add-in needs to open a different page of the task pane after receiving the message, you can use the `window.location.replace` method (or `window.location.href`) as the last line of the handler. The following is an example:</span></span>

```js
function processMessage(arg) {
    // message processing code goes here;
    window.location.replace("/newPage.html");
    // Alternatively ...
    // window.location.href = "/newPage.html";
}
```

<span data-ttu-id="00ecb-198">Para ver um exemplo de um suplemento que faz isso, consulte [Inserir gráficos do Excel usando o Microsoft Graph em um Suplemento do PowerPoint](https://github.com/OfficeDev/PowerPoint-Add-in-Microsoft-Graph-ASPNET-InsertChart).</span><span class="sxs-lookup"><span data-stu-id="00ecb-198">For an example of an add-in that does this, see the [Insert Excel charts using Microsoft Graph in a PowerPoint add-in](https://github.com/OfficeDev/PowerPoint-Add-in-Microsoft-Graph-ASPNET-InsertChart) sample.</span></span>

### <a name="conditional-messaging"></a><span data-ttu-id="00ecb-199">Mensagens condicionais</span><span class="sxs-lookup"><span data-stu-id="00ecb-199">Conditional messaging</span></span>

<span data-ttu-id="00ecb-200">Como você pode enviar várias chamadas `messageParent` a partir da caixa de diálogo, mas tem apenas um manipulador na página host do evento `DialogMessageReceived`, o manipulador tem que usar a lógica condicional para distinguir mensagens diferentes.</span><span class="sxs-lookup"><span data-stu-id="00ecb-200">Because you can send multiple `messageParent` calls from the dialog box, but you have only one handler in the host page for the `DialogMessageReceived` event, the handler must use conditional logic to distinguish different messages.</span></span> <span data-ttu-id="00ecb-201">Por exemplo, se a caixa de diálogo solicitar que um usuário entre em um provedor de identidade como a conta da Microsoft ou o Google, ela enviará o perfil do usuário como uma mensagem.</span><span class="sxs-lookup"><span data-stu-id="00ecb-201">For example, if the dialog box prompts a user to sign in to an identity provider such as Microsoft account or Google, it sends the user's profile as a message.</span></span> <span data-ttu-id="00ecb-202">Se a autenticação falhar, a caixa de diálogo enviará informações de erro à página host, como no exemplo a seguir:</span><span class="sxs-lookup"><span data-stu-id="00ecb-202">If authentication fails, the dialog box sends error information to the host page, as in the following example:</span></span>

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
> - <span data-ttu-id="00ecb-203">A variável `loginSuccess` poderia ser inicializada por meio da leitura da resposta HTTP no provedor de identidade.</span><span class="sxs-lookup"><span data-stu-id="00ecb-203">The `loginSuccess` variable would be initialized by reading the HTTP response from the identity provider.</span></span>
> - <span data-ttu-id="00ecb-p128">A implementação das funções `getProfile` e `getError` não é exibida. Cada uma delas obtém dados de um parâmetro de consulta ou do corpo da resposta HTTP.</span><span class="sxs-lookup"><span data-stu-id="00ecb-p128">The the implementation of the `getProfile` and `getError` functions are not not shown. They each get data from a query parameter or from the body of the HTTP response.</span></span>
> - <span data-ttu-id="00ecb-p129">São enviados objetos anônimos de diferentes tipos se a entrada for bem-sucedida ou não. Ambos têm uma propriedade `messageType`, mas um tem uma propriedade `profile` e o outro tem uma propriedade `error`.</span><span class="sxs-lookup"><span data-stu-id="00ecb-p129">Anonymous objects of different types are sent depending on whether the sign in was successful. Both have a `messageType` property, but one has a `profile` property and the other has an `error` property.</span></span>

<span data-ttu-id="00ecb-p130">O código do manipulador na página host usa o valor da propriedade `messageType` para ramificar como no exemplo a seguir. A função `showUserName` é a mesma do exemplo anterior e a função `showNotification` exibe o erro na interface do usuário da página host.</span><span class="sxs-lookup"><span data-stu-id="00ecb-p130">The handler code in the host page uses the value of the `messageType` property to branch as shown in the following example. Note that the `showUserName` function is the same as in the previous example and `showNotification` function displays the error in the host page's UI.</span></span>

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
> <span data-ttu-id="00ecb-210">A `showNotification` implementação não é exibida no código de exemplo fornecido neste artigo.</span><span class="sxs-lookup"><span data-stu-id="00ecb-210">The `showNotification` implementation is not shown in the sample code provided by this article.</span></span> <span data-ttu-id="00ecb-211">Um exemplo de como você pode implementar essa função dentro do suplemento, confira [Exemplo do suplemento do Office exemplo do diálogo API](https://github.com/OfficeDev/Office-Add-in-Dialog-API-Simple-Example).</span><span class="sxs-lookup"><span data-stu-id="00ecb-211">For an example of how you might implement this function within your add-in, see [Office Add-in Dialog API Example](https://github.com/OfficeDev/Office-Add-in-Dialog-API-Simple-Example).</span></span>

## <a name="pass-information-to-the-dialog-box"></a><span data-ttu-id="00ecb-212">Transmitir informações para a caixa diálogo</span><span class="sxs-lookup"><span data-stu-id="00ecb-212">Pass information to the dialog box</span></span>

<span data-ttu-id="00ecb-213">Seu complemento pode enviar mensagens da página [host](dialog-api-in-office-add-ins.md#open-a-dialog-box-from-a-host-page) para uma caixa de diálogo usando [Dialog.messageChild](/javascript/api/office/office.dialog#messagechild-message-).</span><span class="sxs-lookup"><span data-stu-id="00ecb-213">Your add-in can send messages from the [host page](dialog-api-in-office-add-ins.md#open-a-dialog-box-from-a-host-page) to a dialog box using [Dialog.messageChild](/javascript/api/office/office.dialog#messagechild-message-).</span></span>

### <a name="use-messagechild-from-the-host-page"></a><span data-ttu-id="00ecb-214">Usar `messageChild()` da página host</span><span class="sxs-lookup"><span data-stu-id="00ecb-214">Use `messageChild()` from the host page</span></span>

<span data-ttu-id="00ecb-215">Quando você chama a API de caixa de diálogo do Office para abrir uma caixa de diálogo, um [objeto Dialog](/javascript/api/office/office.dialog) é retornado.</span><span class="sxs-lookup"><span data-stu-id="00ecb-215">When you call the Office dialog API to open a dialog box, a [Dialog](/javascript/api/office/office.dialog) object is returned.</span></span> <span data-ttu-id="00ecb-216">Ele deve ser atribuído a uma variável que tenha um escopo maior do que o [método displayDialogAsync](/javascript/api/office/office.ui#displaydialogasync-startaddress--callback-) porque o objeto será referenciado por outros métodos.</span><span class="sxs-lookup"><span data-stu-id="00ecb-216">It should be assigned to a variable that has greater scope than the [displayDialogAsync](/javascript/api/office/office.ui#displaydialogasync-startaddress--callback-) method because the object will be referenced by other methods.</span></span> <span data-ttu-id="00ecb-217">Veja um exemplo a seguir:</span><span class="sxs-lookup"><span data-stu-id="00ecb-217">The following is an example:</span></span>

```javascript
var dialog;
Office.context.ui.displayDialogAsync('https://myDomain/myDialog.html',
    function (asyncResult) {
        dialog = asyncResult.value;
        dialog.addEventHandler(Office.EventType.DialogMessageReceived, processMessage);
    }
);

function processMessage(arg) {
    dialog.close();

  // message processing code goes here;

}
```

<span data-ttu-id="00ecb-218">Esse objeto tem um método messageChild que envia qualquer cadeia de `Dialog` caracteres, [](/javascript/api/office/office.dialog#messagechild-message-) incluindo dados em cadeia de caracteres, para a caixa de diálogo.</span><span class="sxs-lookup"><span data-stu-id="00ecb-218">This `Dialog` object has a [messageChild](/javascript/api/office/office.dialog#messagechild-message-) method that sends any string, including stringified data, to the dialog box.</span></span> <span data-ttu-id="00ecb-219">Isso gera um `DialogParentMessageReceived` evento na caixa de diálogo.</span><span class="sxs-lookup"><span data-stu-id="00ecb-219">This raises a `DialogParentMessageReceived` event in the dialog box.</span></span> <span data-ttu-id="00ecb-220">Seu código deve manipular esse evento, conforme mostrado na próxima seção.</span><span class="sxs-lookup"><span data-stu-id="00ecb-220">Your code should handle this event, as shown in the next section.</span></span>

<span data-ttu-id="00ecb-221">Considere um cenário no qual a interface do usuário da caixa de diálogo está relacionada à planilha ativa no momento e à posição dessa planilha em relação às outras planilhas.</span><span class="sxs-lookup"><span data-stu-id="00ecb-221">Consider a scenario in which the UI of the dialog is related to the currently active worksheet and that worksheet's position relative to the other worksheets.</span></span> <span data-ttu-id="00ecb-222">No exemplo a seguir, `sheetPropertiesChanged` envia as propriedades da planilha do Excel para a caixa de diálogo.</span><span class="sxs-lookup"><span data-stu-id="00ecb-222">In the following example, `sheetPropertiesChanged` sends Excel worksheet properties to the dialog box.</span></span> <span data-ttu-id="00ecb-223">Nesse caso, a planilha atual é chamada de "Minha Planilha" e é a segunda planilha na pasta de trabalho.</span><span class="sxs-lookup"><span data-stu-id="00ecb-223">In this case, the current worksheet is named "My Sheet" and it's the second sheet in the workbook.</span></span> <span data-ttu-id="00ecb-224">Os dados são encapsulados em um objeto e stringificados para que possam ser passados para `messageChild` .</span><span class="sxs-lookup"><span data-stu-id="00ecb-224">The data is encapsulated in an object and stringified so that it can be passed to `messageChild`.</span></span>

```javascript
function sheetPropertiesChanged() {
    var messageToDialog = JSON.stringify({
                               name: "My Sheet",
                               position: 2
                           });

    dialog.messageChild(messageToDialog);
}
```

### <a name="handle-dialogparentmessagereceived-in-the-dialog-box"></a><span data-ttu-id="00ecb-225">Manipular DialogParentMessageReceived na caixa de diálogo</span><span class="sxs-lookup"><span data-stu-id="00ecb-225">Handle DialogParentMessageReceived in the dialog box</span></span>

<span data-ttu-id="00ecb-226">No JavaScript da caixa de diálogo, registre um manipulador para o evento com o método `DialogParentMessageReceived` [UI.addHandlerAsync.](/javascript/api/office/office.ui#addhandlerasync-eventtype--handler--options--callback-)</span><span class="sxs-lookup"><span data-stu-id="00ecb-226">In the dialog box's JavaScript, register a handler for the `DialogParentMessageReceived` event with the [UI.addHandlerAsync](/javascript/api/office/office.ui#addhandlerasync-eventtype--handler--options--callback-) method.</span></span> <span data-ttu-id="00ecb-227">Isso geralmente é feito nos métodos [Office.onReady ou Office.initialize,](initialize-add-in.md)conforme mostrado a seguir.</span><span class="sxs-lookup"><span data-stu-id="00ecb-227">This is typically done in the [Office.onReady or Office.initialize methods](initialize-add-in.md), as shown in the following.</span></span> <span data-ttu-id="00ecb-228">(Um exemplo mais robusto está abaixo.)</span><span class="sxs-lookup"><span data-stu-id="00ecb-228">(A more robust example is below.)</span></span>

```javascript
Office.onReady()
    .then(function() {
        Office.context.ui.addHandlerAsync(
            Office.EventType.DialogParentMessageReceived,
            onMessageFromParent);
    });
```

<span data-ttu-id="00ecb-229">Em seguida, defina o `onMessageFromParent` manipulador.</span><span class="sxs-lookup"><span data-stu-id="00ecb-229">Then, define the `onMessageFromParent` handler.</span></span> <span data-ttu-id="00ecb-230">O código a seguir continua o exemplo da seção anterior.</span><span class="sxs-lookup"><span data-stu-id="00ecb-230">The following code continues the example from the preceding section.</span></span> <span data-ttu-id="00ecb-231">Observe que o Office passa um argumento para o manipulador e que a propriedade do objeto de argumento contém a `message` cadeia de caracteres da página host.</span><span class="sxs-lookup"><span data-stu-id="00ecb-231">Note that Office passes an argument to the handler and that the `message` property of the argument object contains the string from the host page.</span></span> <span data-ttu-id="00ecb-232">Neste exemplo, a mensagem é revertida para um objeto e jQuery é usado para definir o título superior da caixa de diálogo para corresponder ao novo nome da planilha.</span><span class="sxs-lookup"><span data-stu-id="00ecb-232">In this example, the message is reconverted to an object and jQuery is used to set the top heading of the dialog to match the new worksheet name.</span></span>

```javascript
function onMessageFromParent(event) {
    var messageFromParent = JSON.parse(event.message);
    $('h1').text(messageFromParent.name);
}
```

<span data-ttu-id="00ecb-233">É uma prática adequada verificar se o manipulador está registrado corretamente.</span><span class="sxs-lookup"><span data-stu-id="00ecb-233">It is a best practice to verify that your handler is properly registered.</span></span> <span data-ttu-id="00ecb-234">Você pode fazer isso passando um retorno de chamada para o `addHandlerAsync` método.</span><span class="sxs-lookup"><span data-stu-id="00ecb-234">You can do this by passing a callback to the `addHandlerAsync` method.</span></span> <span data-ttu-id="00ecb-235">Isso é executado quando a tentativa de registrar o manipulador é concluída.</span><span class="sxs-lookup"><span data-stu-id="00ecb-235">This runs when the attempt to register the handler completes.</span></span> <span data-ttu-id="00ecb-236">Use o manipulador para registrar ou mostrar um erro se o manipulador não tiver sido registrado com êxito.</span><span class="sxs-lookup"><span data-stu-id="00ecb-236">Use the handler to log or show an error if the handler was not successfully registered.</span></span> <span data-ttu-id="00ecb-237">Apresentamos um exemplo a seguir.</span><span class="sxs-lookup"><span data-stu-id="00ecb-237">The following is an example.</span></span> <span data-ttu-id="00ecb-238">Observe que `reportError` é uma função, não definida aqui, que registra ou exibe o erro.</span><span class="sxs-lookup"><span data-stu-id="00ecb-238">Note that `reportError` is a function, not defined here, that logs or displays the error.</span></span>

```javascript
Office.onReady()
    .then(function() {
        Office.context.ui.addHandlerAsync(
            Office.EventType.DialogParentMessageReceived,
            onMessageFromParent,
            onRegisterMessageComplete);
    });

function onRegisterMessageComplete(asyncResult) {
    if (asyncResult.status !== Office.AsyncResultStatus.Succeeded) {
        reportError(asyncResult.error.message);
    }
}
```

### <a name="conditional-messaging-from-parent-page-to-dialog-box"></a><span data-ttu-id="00ecb-239">Mensagens condicionais da página pai para a caixa de diálogo</span><span class="sxs-lookup"><span data-stu-id="00ecb-239">Conditional messaging from parent page to dialog box</span></span>

<span data-ttu-id="00ecb-240">Como você pode fazer várias chamadas da página host, mas tem apenas um manipulador na caixa de diálogo para o evento, o manipulador deve usar lógica condicional para distinguir `messageChild` `DialogParentMessageReceived` mensagens diferentes.</span><span class="sxs-lookup"><span data-stu-id="00ecb-240">Because you can make multiple `messageChild` calls from the host page, but you have only one handler in the dialog box for the `DialogParentMessageReceived` event, the handler must use conditional logic to distinguish different messages.</span></span> <span data-ttu-id="00ecb-241">Você pode fazer isso de uma maneira que seja precisamente paralela a como estruturaria as mensagens condicionais quando a caixa de diálogo está enviando uma mensagem para a página host, conforme descrito em mensagens [condicionais.](#conditional-messaging)</span><span class="sxs-lookup"><span data-stu-id="00ecb-241">You can do this in a way that is precisely parallel to how you would structure conditional messaging when the dialog box is sending a message to the host page as described in [Conditional messaging](#conditional-messaging).</span></span>

> [!NOTE]
> <span data-ttu-id="00ecb-242">Em algumas situações, a API, que faz parte do conjunto de requisitos `messageChild` [DialogApi 1.2,](../reference/requirement-sets/dialog-api-requirement-sets.md)pode não ter suporte.</span><span class="sxs-lookup"><span data-stu-id="00ecb-242">In some situations, the `messageChild` API, which is a part of the [DialogApi 1.2 requirement set](../reference/requirement-sets/dialog-api-requirement-sets.md),  may not be supported.</span></span> <span data-ttu-id="00ecb-243">Algumas maneiras alternativas para mensagens pai-para-caixa de diálogo são descritas de maneiras alternativas de passar mensagens para uma caixa de diálogo [de sua página host](parent-to-dialog.md).</span><span class="sxs-lookup"><span data-stu-id="00ecb-243">Some alternative ways for parent-to-dialog-box messaging are described in [Alternative ways of passing messages to a dialog box from its host page](parent-to-dialog.md).</span></span>

> [!IMPORTANT]
> <span data-ttu-id="00ecb-244">O [conjunto de requisitos DialogApi 1.2](../reference/requirement-sets/dialog-api-requirement-sets.md) não pode ser especificado na seção `<Requirements>` de um manifesto de um complemento.</span><span class="sxs-lookup"><span data-stu-id="00ecb-244">The [DialogApi 1.2 requirement set](../reference/requirement-sets/dialog-api-requirement-sets.md) cannot be specified in the `<Requirements>` section of an add-in manifest.</span></span> <span data-ttu-id="00ecb-245">Você terá que verificar se há suporte para DialogApi 1.2 no tempo de execução usando o [método isSetSupported.](specify-office-hosts-and-api-requirements.md#use-runtime-checks-in-your-javascript-code)</span><span class="sxs-lookup"><span data-stu-id="00ecb-245">You will have to check for support for DialogApi 1.2 at runtime using the [isSetSupported](specify-office-hosts-and-api-requirements.md#use-runtime-checks-in-your-javascript-code) method.</span></span> <span data-ttu-id="00ecb-246">O suporte para requisitos de manifesto está em desenvolvimento.</span><span class="sxs-lookup"><span data-stu-id="00ecb-246">Support for manifest requirements is under development.</span></span>

## <a name="closing-the-dialog-box"></a><span data-ttu-id="00ecb-247">Feche a caixa de diálogo</span><span class="sxs-lookup"><span data-stu-id="00ecb-247">Closing the dialog box</span></span>

<span data-ttu-id="00ecb-p141">Você pode implementar um botão na caixa de diálogo para fechá-la. Para fazer isso, o manipulador de eventos de clique do botão deve usar `messageParent` para informar a página host em que o botão foi clicado. Apresentamos um exemplo a seguir:</span><span class="sxs-lookup"><span data-stu-id="00ecb-p141">You can implement a button in the dialog box that will close it. To do this, the click event handler for the button should use `messageParent` to tell the host page that the button has been clicked. The following is an example:</span></span>

```js
function closeButtonClick() {
    var messageObject = {messageType: "dialogClosed"};
    var jsonMessage = JSON.stringify(messageObject);
    Office.context.ui.messageParent(jsonMessage);
}
```

<span data-ttu-id="00ecb-251">O manipulador de página host de `DialogMessageReceived` poderia chamar `dialog.close`, como neste exemplo.</span><span class="sxs-lookup"><span data-stu-id="00ecb-251">The host page handler for `DialogMessageReceived` would call `dialog.close`, as in this example.</span></span> <span data-ttu-id="00ecb-252">(Veja exemplos anteriores que mostram como o objeto `dialog` é inicializado.)</span><span class="sxs-lookup"><span data-stu-id="00ecb-252">(See previous examples that show how the `dialog` object is initialized.)</span></span>

```js
function processMessage(arg) {
    var messageFromDialog = JSON.parse(arg.message);
    if (messageFromDialog.messageType === "dialogClosed") {
       dialog.close();
    }
}
```

<span data-ttu-id="00ecb-253">Mesmo quando você não tem sua própria interface de usuário de diálogo de fechar, um usuário final pode fechar a caixa de diálogo escolhendo a opção **X** no canto superior direito.</span><span class="sxs-lookup"><span data-stu-id="00ecb-253">Even when you don't have your own close-dialog UI, an end user can close the dialog box by choosing the **X** in the upper-right corner.</span></span> <span data-ttu-id="00ecb-254">Essa ação aciona o evento `DialogEventReceived`.</span><span class="sxs-lookup"><span data-stu-id="00ecb-254">This action triggers the `DialogEventReceived` event.</span></span> <span data-ttu-id="00ecb-255">Se seu painel do host precisar saber quando isso acontece, ele deverá declarar um manipulador para esse evento.</span><span class="sxs-lookup"><span data-stu-id="00ecb-255">If your host pane needs to know when this happens, it should declare a handler for this event.</span></span> <span data-ttu-id="00ecb-256">Confira a seção [Erros e eventos na caixa de diálogo](dialog-handle-errors-events.md#errors-and-events-in-the-dialog-box) para ver os detalhes.</span><span class="sxs-lookup"><span data-stu-id="00ecb-256">See the section [Errors and events in the dialog box](dialog-handle-errors-events.md#errors-and-events-in-the-dialog-box) for details.</span></span>

## <a name="advanced-topics-and-special-scenarios"></a><span data-ttu-id="00ecb-257">Tópicos avançados e cenários especiais</span><span class="sxs-lookup"><span data-stu-id="00ecb-257">Advanced topics and special scenarios</span></span>

### <a name="use-the-dialog-api-to-show-a-video"></a><span data-ttu-id="00ecb-258">Use a API de Caixa de Diálogo para exibir um vídeo</span><span class="sxs-lookup"><span data-stu-id="00ecb-258">Use the Dialog API to show a video</span></span>

<span data-ttu-id="00ecb-259">Confira [use a caixa de diálogo do Office para mostrar um vídeo](dialog-video.md).</span><span class="sxs-lookup"><span data-stu-id="00ecb-259">See [Use the Office dialog box to show a video](dialog-video.md).</span></span>

### <a name="use-the-dialog-apis-in-an-authentication-flow"></a><span data-ttu-id="00ecb-260">Use as APIs de Caixa de Diálogo em um fluxo de autenticação</span><span class="sxs-lookup"><span data-stu-id="00ecb-260">Use the Dialog APIs in an authentication flow</span></span>

<span data-ttu-id="00ecb-261">Confira[Autenticar com a API da Caixa de Diálogo do Office](auth-with-office-dialog-api.md).</span><span class="sxs-lookup"><span data-stu-id="00ecb-261">See [Authenticate with the Office dialog API](auth-with-office-dialog-api.md).</span></span>

### <a name="using-the-office-dialog-api-with-single-page-applications-and-client-side-routing"></a><span data-ttu-id="00ecb-262">Usar a API de Caixa de diálogo do Office com aplicativos de página única e roteamento do lado do cliente</span><span class="sxs-lookup"><span data-stu-id="00ecb-262">Using the Office dialog API with single-page applications and client-side routing</span></span>

<span data-ttu-id="00ecb-263">SPAs e o roteamento do lado do cliente devem ser manuseados com cuidado ao usar a API de diálogo do Office.</span><span class="sxs-lookup"><span data-stu-id="00ecb-263">SPAs and client-side routing need to be handled with care when you are using the Office dialog API.</span></span> <span data-ttu-id="00ecb-264">Confira [práticas recomendadas para usar o Office Dialog API em um SPA](dialog-best-practices.md#best-practices-for-using-the-office-dialog-api-in-an-spa).</span><span class="sxs-lookup"><span data-stu-id="00ecb-264">Please see [Best practices for using the Office dialog API in an SPA](dialog-best-practices.md#best-practices-for-using-the-office-dialog-api-in-an-spa).</span></span>

### <a name="error-and-event-handling"></a><span data-ttu-id="00ecb-265">Manipulação de erros e eventos</span><span class="sxs-lookup"><span data-stu-id="00ecb-265">Error and event handling</span></span>

<span data-ttu-id="00ecb-266">Confira [Manipulando erros e eventos na caixa de diálogo do Office](dialog-handle-errors-events.md).</span><span class="sxs-lookup"><span data-stu-id="00ecb-266">See [Handling errors and events in the Office dialog box](dialog-handle-errors-events.md).</span></span>

## <a name="next-steps"></a><span data-ttu-id="00ecb-267">Próximas etapas</span><span class="sxs-lookup"><span data-stu-id="00ecb-267">Next steps</span></span>

<span data-ttu-id="00ecb-268">Saiba mais sobre as armadilhas e as práticas recomendadas para a API de diálogo do Office em [práticas recomendadas e regras para a API do Office Dialog](dialog-best-practices.md).</span><span class="sxs-lookup"><span data-stu-id="00ecb-268">Learn about gotchas and best practices for the Office dialog API in [Best practices and rules for the Office dialog API](dialog-best-practices.md).</span></span>

## <a name="samples"></a><span data-ttu-id="00ecb-269">Exemplos</span><span class="sxs-lookup"><span data-stu-id="00ecb-269">Samples</span></span>

<span data-ttu-id="00ecb-270">Todos os exemplos a seguir `displayDialogAsync` usam.</span><span class="sxs-lookup"><span data-stu-id="00ecb-270">All of the following samples use `displayDialogAsync`.</span></span> <span data-ttu-id="00ecb-271">Alguns têm servidores baseados em NodeJS e outros têm servidores ASP.NET/IIS-based, mas a lógica de uso do método é a mesma, independentemente de como o lado do servidor do add-in é implementado.</span><span class="sxs-lookup"><span data-stu-id="00ecb-271">Some have NodeJS-based servers and others have ASP.NET/IIS-based servers, but the logic of using the method is the same regardless of how the server-side of the add-in is implemented.</span></span>

<span data-ttu-id="00ecb-272">**Noções básicas:**</span><span class="sxs-lookup"><span data-stu-id="00ecb-272">**Basics:**</span></span>

- [<span data-ttu-id="00ecb-273">Exemplo da API da caixa de diálogo do suplemento do Office</span><span class="sxs-lookup"><span data-stu-id="00ecb-273">Office Add-in Dialog API Example</span></span>](https://github.com/OfficeDev/Office-Add-in-Dialog-API-Simple-Example)
- [<span data-ttu-id="00ecb-274">Conteúdo de treinamento/criação de complementos (vários exemplos)</span><span class="sxs-lookup"><span data-stu-id="00ecb-274">Training Content / Building Add-ins (several samples)</span></span>](https://github.com/OfficeDev/TrainingContent/tree/2db14a16774e1539a3eebae7dada4798142b8493/OfficeAddin)

<span data-ttu-id="00ecb-275">**Exemplos mais complexos:**</span><span class="sxs-lookup"><span data-stu-id="00ecb-275">**More complex samples:**</span></span>

- [<span data-ttu-id="00ecb-276">ASPNET do Microsoft Graph para o complemento do Office</span><span class="sxs-lookup"><span data-stu-id="00ecb-276">Office Add-in Microsoft Graph ASPNET</span></span>](https://github.com/OfficeDev/PnP-OfficeAddins/tree/master/Samples/auth/Office-Add-in-Microsoft-Graph-ASPNET)
- [<span data-ttu-id="00ecb-277">Suplemento do Office Microsoft Graph React</span><span class="sxs-lookup"><span data-stu-id="00ecb-277">Office Add-in Microsoft Graph React</span></span>](https://github.com/OfficeDev/PnP-OfficeAddins/tree/master/Samples/auth/Office-Add-in-Microsoft-Graph-React)
- [<span data-ttu-id="00ecb-278">SSO do NodeJS do Suplemento do Office</span><span class="sxs-lookup"><span data-stu-id="00ecb-278">Office Add-in NodeJS SSO</span></span>](https://github.com/OfficeDev/Office-Add-in-NodeJS-SSO)
- [<span data-ttu-id="00ecb-279">Office Add-in ASPNET SSO</span><span class="sxs-lookup"><span data-stu-id="00ecb-279">Office Add-in ASPNET SSO</span></span>](https://github.com/OfficeDev/Office-Add-in-ASPNET-SSO)
- [<span data-ttu-id="00ecb-280">Exemplo de monetização SAAS de complemento do Office</span><span class="sxs-lookup"><span data-stu-id="00ecb-280">Office Add-in SAAS Monetization Sample</span></span>](https://github.com/OfficeDev/office-add-in-saas-monetization-sample)
- [<span data-ttu-id="00ecb-281">Outlook Add-in Microsoft Graph ASPNET</span><span class="sxs-lookup"><span data-stu-id="00ecb-281">Outlook Add-in Microsoft Graph ASPNET</span></span>](https://github.com/OfficeDev/PnP-OfficeAddins/tree/master/Samples/auth/Outlook-Add-in-Microsoft-Graph-ASPNET)
- [<span data-ttu-id="00ecb-282">Outlook Add-in SSO</span><span class="sxs-lookup"><span data-stu-id="00ecb-282">Outlook Add-in SSO</span></span>](https://github.com/OfficeDev/Outlook-Add-in-SSO)
- [<span data-ttu-id="00ecb-283">Visualizador de Tokens de Complementos do Outlook</span><span class="sxs-lookup"><span data-stu-id="00ecb-283">Outlook Add-in Token Viewer</span></span>](https://github.com/OfficeDev/Outlook-Add-In-Token-Viewer)
- [<span data-ttu-id="00ecb-284">Mensagem Acionável do Complemento do Outlook</span><span class="sxs-lookup"><span data-stu-id="00ecb-284">Outlook Add-in Actionable Message</span></span>](https://github.com/OfficeDev/Outlook-Add-In-Actionable-Message)
- [<span data-ttu-id="00ecb-285">Compartilhamento de Complementos do Outlook com o OneDrive</span><span class="sxs-lookup"><span data-stu-id="00ecb-285">Outlook Add-in Sharing to OneDrive</span></span>](https://github.com/OfficeDev/Outlook-Add-in-Sharing-to-OneDrive)
- [<span data-ttu-id="00ecb-286">PowerPoint Add-in Microsoft Graph ASPNET InsertChart</span><span class="sxs-lookup"><span data-stu-id="00ecb-286">PowerPoint Add-in Microsoft Graph ASPNET InsertChart</span></span>](https://github.com/OfficeDev/PowerPoint-Add-in-Microsoft-Graph-ASPNET-InsertChart)
- [<span data-ttu-id="00ecb-287">Cenário de Tempo de Execução Compartilhado do Excel</span><span class="sxs-lookup"><span data-stu-id="00ecb-287">Excel Shared Runtime Scenario</span></span>](https://github.com/OfficeDev/PnP-OfficeAddins/tree/900b5769bca9bbcff79d6cd6106d9fcc55c70d5a/Samples/excel-shared-runtime-scenario)
- [<span data-ttu-id="00ecb-288">Guias rápidos asPNET de um complemento do Excel</span><span class="sxs-lookup"><span data-stu-id="00ecb-288">Excel Add-in ASPNET QuickBooks</span></span>](https://github.com/OfficeDev/Excel-Add-in-ASPNET-QuickBooks)
- [<span data-ttu-id="00ecb-289">Word Add-in JS Redact</span><span class="sxs-lookup"><span data-stu-id="00ecb-289">Word Add-in JS Redact</span></span>](https://github.com/OfficeDev/Word-Add-in-JS-Redact)
- [<span data-ttu-id="00ecb-290">Word Add-in JS SpecKit</span><span class="sxs-lookup"><span data-stu-id="00ecb-290">Word Add-in JS SpecKit</span></span>](https://github.com/OfficeDev/Word-Add-in-JS-SpecKit)
- [<span data-ttu-id="00ecb-291">OAuth do cliente AngularJS do Word Add-in</span><span class="sxs-lookup"><span data-stu-id="00ecb-291">Word Add-in AngularJS Client OAuth</span></span>](https://github.com/OfficeDev/Word-Add-in-AngularJS-Client-OAuth)
- [<span data-ttu-id="00ecb-292">Suplemento do Office Auth0</span><span class="sxs-lookup"><span data-stu-id="00ecb-292">Office Add-in Auth0</span></span>](https://github.com/OfficeDev/Office-Add-in-Auth0)
- [<span data-ttu-id="00ecb-293">Office Add-in OAuth.io</span><span class="sxs-lookup"><span data-stu-id="00ecb-293">Office Add-in OAuth.io</span></span>](https://github.com/OfficeDev/Office-Add-in-OAuth.io)
- [<span data-ttu-id="00ecb-294">Código de padrões de design da UX de um complemento do Office</span><span class="sxs-lookup"><span data-stu-id="00ecb-294">Office Add-in UX Design Patterns Code</span></span>](https://github.com/OfficeDev/Office-Add-in-UX-Design-Patterns-Code)
