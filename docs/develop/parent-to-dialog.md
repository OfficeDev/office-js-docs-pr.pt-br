---
title: Passando dados e mensagens para uma caixa de diálogo da página host
description: Saiba como transmitir dados para uma caixa de diálogo da página host usando as APIs messageChild e DialogParentMessageReceived
ms.date: 03/11/2020
localization_priority: Normal
ms.openlocfilehash: 03d89a2e5ffb9060edb25dd8e0c3c71c0dd274eb
ms.sourcegitcommit: 153576b1efd0234c6252433e22db213238573534
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 03/07/2020
ms.locfileid: "42561867"
---
# <a name="passing-data-and-messages-to-a-dialog-box-from-its-host-page-preview"></a><span data-ttu-id="eba88-103">Passando dados e mensagens para uma caixa de diálogo da página de host (visualização)</span><span class="sxs-lookup"><span data-stu-id="eba88-103">Passing data and messages to a dialog box from its host page (preview)</span></span>

<span data-ttu-id="eba88-104">O suplemento pode enviar mensagens da [página host](dialog-api-in-office-add-ins.md#open-a-dialog-box-from-a-host-page) para uma caixa de diálogo usando o método [MessageChild](/javascript/api/office/office.dialog#messagechild-message-) do objeto [Dialog](/javascript/api/office/office.dialog) .</span><span class="sxs-lookup"><span data-stu-id="eba88-104">Your add-in can send messages from the [host page](dialog-api-in-office-add-ins.md#open-a-dialog-box-from-a-host-page) to a dialog box using the [messageChild](/javascript/api/office/office.dialog#messagechild-message-) method of the [Dialog](/javascript/api/office/office.dialog) object.</span></span>

> [!Important]
>
> - <span data-ttu-id="eba88-105">As APIs descritas neste artigo estão em visualização.</span><span class="sxs-lookup"><span data-stu-id="eba88-105">The APIs described in this article are in preview.</span></span> <span data-ttu-id="eba88-106">Eles estão disponíveis para os desenvolvedores de experimentação; Mas não deve ser usado em um suplemento de produção.</span><span class="sxs-lookup"><span data-stu-id="eba88-106">They are available to developers for experimentation; but should not be used in a production add-in.</span></span> <span data-ttu-id="eba88-107">Até que esta API seja liberada, use as técnicas descritas em [passar informações para a caixa de diálogo](dialog-api-in-office-add-ins.md#pass-information-to-the-dialog-box) para suplementos de produção.</span><span class="sxs-lookup"><span data-stu-id="eba88-107">Until this API is released, use the techniques described in [Pass information to the dialog box](dialog-api-in-office-add-ins.md#pass-information-to-the-dialog-box) for production add-ins.</span></span>
> - <span data-ttu-id="eba88-108">As APIs descritas neste artigo exigem o Office 365 (a versão de assinatura do Office).</span><span class="sxs-lookup"><span data-stu-id="eba88-108">The APIs described in this article require Office 365 (the subscription version of Office).</span></span> <span data-ttu-id="eba88-109">Você deve usar o build e a versão mensal mais recente do canal Insiders.</span><span class="sxs-lookup"><span data-stu-id="eba88-109">You should use the latest monthly version and build from the Insiders channel.</span></span> <span data-ttu-id="eba88-110">É necessário ingressar no programa Office Insider para obter essa versão.</span><span class="sxs-lookup"><span data-stu-id="eba88-110">You need to be an Office Insider to get this version.</span></span> <span data-ttu-id="eba88-111">Para saber mais, confira a página [Seja um Office Insider](https://products.office.com/office-insider?tab=tab-1).</span><span class="sxs-lookup"><span data-stu-id="eba88-111">For more information, see [Be an Office Insider](https://products.office.com/office-insider?tab=tab-1).</span></span> <span data-ttu-id="eba88-112">Observe que, quando uma compilação é graduada para o canal semestral de produção, o suporte para recursos de visualização é desativado para essa compilação.</span><span class="sxs-lookup"><span data-stu-id="eba88-112">Please note that when a build graduates to the production semi-annual channel, support for preview features is turned off for that build.</span></span>
> - <span data-ttu-id="eba88-113">Na fase inicial da visualização, as APIs têm suporte no Excel, PowerPoint e Word; Mas não no Outlook.</span><span class="sxs-lookup"><span data-stu-id="eba88-113">In the initial stage of the preview, the APIs are supported in Excel, PowerPoint, and Word; but not in Outlook.</span></span>
>
> [!INCLUDE [Information about using preview APIs](../includes/using-preview-apis.md)]

## <a name="use-messagechild-from-the-host-page"></a><span data-ttu-id="eba88-114">Usar `messageChild()` na página host</span><span class="sxs-lookup"><span data-stu-id="eba88-114">Use `messageChild()` from the host page</span></span>

<span data-ttu-id="eba88-115">Quando você chama a API de diálogo do Office para abrir uma caixa de diálogo, um objeto [Dialog](/javascript/api/office/office.dialog) é retornado.</span><span class="sxs-lookup"><span data-stu-id="eba88-115">When you call the Office dialog API to open a dialog box, a [Dialog](/javascript/api/office/office.dialog) object is returned.</span></span> <span data-ttu-id="eba88-116">Ele deve ser atribuído a uma variável, que geralmente tem escopo maior do que o método [displayDialogAsync](/javascript/api/office/office.ui#displaydialogasync-startaddress--callback-) porque o objeto será referenciado por outros métodos.</span><span class="sxs-lookup"><span data-stu-id="eba88-116">It should be assigned to a variable, which typically has greater scope than the [displayDialogAsync](/javascript/api/office/office.ui#displaydialogasync-startaddress--callback-) method because the object will be referenced by other methods.</span></span> <span data-ttu-id="eba88-117">Este é um exemplo:</span><span class="sxs-lookup"><span data-stu-id="eba88-117">The following is an example:</span></span>

```javascript
var dialog;
Office.context.ui.displayDialogAsync('https://myDomain/myDialog.html',
    function (asyncResult) {
        dialog = asyncResult.value;
        dialog.addEventHandler(Office.EventType.DialogMessageReceived, processMessage);
    }
);

function processMessage(arg) {
    dialog.close();

  // message processing code goes here;

}
```

<span data-ttu-id="eba88-118">Este `Dialog` objeto tem um método [messageChild](/javascript/api/office/office.dialog#messagechild-message-) que envia qualquer cadeia de caracteres ou em formato dados para a caixa de diálogo.</span><span class="sxs-lookup"><span data-stu-id="eba88-118">This `Dialog` object has a [messageChild](/javascript/api/office/office.dialog#messagechild-message-) method that sends any string, or stringified data, to the dialog box.</span></span> <span data-ttu-id="eba88-119">Isso gera um `DialogParentMessageReceived` evento na caixa de diálogo.</span><span class="sxs-lookup"><span data-stu-id="eba88-119">This raises a `DialogParentMessageReceived` event in the dialog box.</span></span> <span data-ttu-id="eba88-120">O código deve lidar com esse evento, conforme mostrado na próxima seção.</span><span class="sxs-lookup"><span data-stu-id="eba88-120">Your code should handle this event, as shown in the next section.</span></span>

<span data-ttu-id="eba88-121">Considere um cenário em que a interface do usuário da caixa de diálogo deve se correlacionar com a planilha ativa no momento e a posição dessa planilha em relação às outras planilhas.</span><span class="sxs-lookup"><span data-stu-id="eba88-121">Consider a scenario in which the UI of the dialog should correlate with the currently active worksheet and that worksheet's position relative to the other worksheets.</span></span> <span data-ttu-id="eba88-122">No exemplo a seguir, `sheetPropertiesChanged` envia as propriedades de planilha do Excel para a caixa de diálogo.</span><span class="sxs-lookup"><span data-stu-id="eba88-122">In the following example, `sheetPropertiesChanged` sends Excel worksheet properties to the dialog box.</span></span> <span data-ttu-id="eba88-123">Nesse caso, a planilha atual é chamada "minha planilha" e é a 2ª folha na pasta de trabalho.</span><span class="sxs-lookup"><span data-stu-id="eba88-123">In this case the current worksheet is named "My Sheet" and it is the 2nd sheet in the workbook.</span></span> <span data-ttu-id="eba88-124">Os dados são encapsulados em um objeto que é em formato para que ele possa ser passado para `messageChild`.</span><span class="sxs-lookup"><span data-stu-id="eba88-124">The data is encapsulated in an object which is stringified so that it can be passed to `messageChild`.</span></span>

```javascript
function sheetPropertiesChanged() {
    var messageToDialog = JSON.stringify({
                               name: "My Sheet",
                               position: 2
                           });

    dialog.messageChild(messageToDialog);
}
```

## <a name="handle-dialogparentmessagereceived-in-the-dialog-box"></a><span data-ttu-id="eba88-125">Manipular DialogParentMessageReceived na caixa de diálogo</span><span class="sxs-lookup"><span data-stu-id="eba88-125">Handle DialogParentMessageReceived in the dialog box</span></span>

<span data-ttu-id="eba88-126">No JavaScript da caixa de diálogo, registre um manipulador para o `DialogParentMessageReceived` evento com o método [UI. addHandlerAsync](/javascript/api/office/office.ui#addhandlerasync-eventtype--handler--options--callback-) .</span><span class="sxs-lookup"><span data-stu-id="eba88-126">In the dialog box's JavaScript, register a handler for the `DialogParentMessageReceived` event with the [UI.addHandlerAsync](/javascript/api/office/office.ui#addhandlerasync-eventtype--handler--options--callback-) method.</span></span> <span data-ttu-id="eba88-127">Isso geralmente é feito nos [métodos Office. onReady ou Office. Initialize](initialize-add-in.md).</span><span class="sxs-lookup"><span data-stu-id="eba88-127">This is typically done in the [Office.onReady or Office.initialize methods](initialize-add-in.md).</span></span> <span data-ttu-id="eba88-128">Este é um exemplo:</span><span class="sxs-lookup"><span data-stu-id="eba88-128">The following is an example:</span></span>

```javascript
Office.onReady()
    .then(function() {
        Office.context.ui.addHandlerAsync(
            Office.EventType.DialogParentMessageReceived,
            onMessageFromParent);
    });
```

<span data-ttu-id="eba88-129">Em seguida, defina `onMessageFromParent` o manipulador.</span><span class="sxs-lookup"><span data-stu-id="eba88-129">Then, define the `onMessageFromParent` handler.</span></span> <span data-ttu-id="eba88-130">O código a seguir continua o exemplo da seção anterior.</span><span class="sxs-lookup"><span data-stu-id="eba88-130">The following code continues the example from the preceding section.</span></span> <span data-ttu-id="eba88-131">Observe que o Office passa um argumento para o manipulador e que `message` a propriedade do objeto Argument contém a cadeia de caracteres da página host.</span><span class="sxs-lookup"><span data-stu-id="eba88-131">Note that Office passes an argument to the handler and that the `message` property of argument object contains the string from the host page.</span></span> <span data-ttu-id="eba88-132">Neste exemplo, a mensagem é convertida para um objeto e o jQuery é usado para definir o título superior da caixa de diálogo para corresponder ao novo nome da planilha.</span><span class="sxs-lookup"><span data-stu-id="eba88-132">In this example, the message is reconverted to an object and jQuery is used to set the top heading of the dialog to match the new worksheet name.</span></span>

```javascript
function onMessageFromParent(event) {
    var messageFromParent = JSON.parse(event.message);
    $('h1').text(messageFromParent.name);
}
```

<span data-ttu-id="eba88-133">É uma prática recomendada verificar se o manipulador está registrado corretamente.</span><span class="sxs-lookup"><span data-stu-id="eba88-133">It is a best practice to verify that your handler is properly registered.</span></span> <span data-ttu-id="eba88-134">Você pode fazer isso passando um retorno de chamada para `addHandlerAsync` o método que é executado quando a tentativa de registrar o manipulador é concluída.</span><span class="sxs-lookup"><span data-stu-id="eba88-134">You can do this by passing a callback to the `addHandlerAsync` method that runs when the attempt to register the handler completes.</span></span> <span data-ttu-id="eba88-135">Use o manipulador para registrar ou mostrar um erro se o manipulador não tiver sido registrado com êxito.</span><span class="sxs-lookup"><span data-stu-id="eba88-135">Use the handler to log or show an error if the handler was not successfully registered.</span></span> <span data-ttu-id="eba88-136">Apresentamos um exemplo a seguir.</span><span class="sxs-lookup"><span data-stu-id="eba88-136">The following is an example.</span></span> <span data-ttu-id="eba88-137">Observe que `reportError` é uma função, não definida aqui, que registra ou exibe o erro.</span><span class="sxs-lookup"><span data-stu-id="eba88-137">Note that `reportError` is a function, not defined here, that logs or displays the error.</span></span>

```javascript
Office.onReady()
    .then(function() {
        Office.context.ui.addHandlerAsync(
            Office.EventType.DialogParentMessageReceived,
            onMessageFromParent,
            onRegisterMessageComplete);
    });

function onRegisterMessageComplete(asyncResult) {
    if (asyncResult.status !== Office.AsyncResultStatus.Succeeded) {
        reportError(asyncResult.error.message);
    }
}
```

## <a name="conditional-messaging"></a><span data-ttu-id="eba88-138">Mensagens condicionais</span><span class="sxs-lookup"><span data-stu-id="eba88-138">Conditional messaging</span></span>

<span data-ttu-id="eba88-139">Como você pode fazer várias `messageChild` chamadas a partir da página host, mas tem apenas um manipulador na caixa de diálogo para o `DialogParentMessageReceived` evento, o manipulador deve usar a lógica condicional para distinguir mensagens diferentes.</span><span class="sxs-lookup"><span data-stu-id="eba88-139">Because you can make multiple `messageChild` calls from the host page, but you have only one handler in the dialog box for the `DialogParentMessageReceived` event, the handler must use conditional logic to distinguish different messages.</span></span> <span data-ttu-id="eba88-140">Você pode fazer isso de uma maneira que seja precisamente paralela à forma como você estruturaria mensagens condicionais quando a caixa de diálogo estiver enviando uma mensagem para a página host, conforme descrito em [mensagens condicionais](dialog-api-in-office-add-ins.md#conditional-messaging).</span><span class="sxs-lookup"><span data-stu-id="eba88-140">You can do this in a way that is precisely parallel to how you would structure conditional messaging when the dialog box is sending a message to the host page as described in [Conditional messaging](dialog-api-in-office-add-ins.md#conditional-messaging).</span></span>
