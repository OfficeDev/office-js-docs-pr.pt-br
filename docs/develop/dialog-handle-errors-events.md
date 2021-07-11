---
title: Manipulando erros e eventos na caixa de diálogo do Office
description: Descreve como interceptar e manipular erros ao abrir e usar Office caixa de diálogo
ms.date: 01/29/2020
localization_priority: Normal
ms.openlocfilehash: be1fb8bcd30b47ac6399657d928d3cad7f857f39
ms.sourcegitcommit: 883f71d395b19ccfc6874a0d5942a7016eb49e2c
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 07/09/2021
ms.locfileid: "53349893"
---
# <a name="handling-errors-and-events-in-the-office-dialog-box"></a><span data-ttu-id="5c9c3-103">Manipulando erros e eventos na caixa de diálogo do Office</span><span class="sxs-lookup"><span data-stu-id="5c9c3-103">Handling errors and events in the Office dialog box</span></span>

<span data-ttu-id="5c9c3-104">Este artigo descreve como interceptar e manipular erros ao abrir a caixa de diálogo e erros que ocorrem dentro da caixa de diálogo.</span><span class="sxs-lookup"><span data-stu-id="5c9c3-104">This article describes how to trap and handle errors when opening the dialog box and errors that happen inside the dialog box.</span></span>

> [!NOTE]
> <span data-ttu-id="5c9c3-105">Este artigo pressupõe que você está familiarizado com as noções básicas de uso da API de diálogo Office conforme descrito em Usar Office API de diálogo Office em seus [Office Add-ins](dialog-api-in-office-add-ins.md)do Office .</span><span class="sxs-lookup"><span data-stu-id="5c9c3-105">This article presupposes that you are familiar with the basics of using the Office dialog API as described in [Use the Office dialog API in your Office Add-ins](dialog-api-in-office-add-ins.md).</span></span>
> 
> <span data-ttu-id="5c9c3-106">Consulte também [Práticas recomendadas e regras para a API Office caixa de diálogo](dialog-best-practices.md).</span><span class="sxs-lookup"><span data-stu-id="5c9c3-106">See also [Best practices and rules for the Office dialog API](dialog-best-practices.md).</span></span>

<span data-ttu-id="5c9c3-107">Seu código deve manipular duas categorias de eventos:</span><span class="sxs-lookup"><span data-stu-id="5c9c3-107">Your code should handle two categories of events:</span></span>

- <span data-ttu-id="5c9c3-108">Erros retornados pela chamada de `displayDialogAsync` porque não foi possível criar a caixa de diálogo.</span><span class="sxs-lookup"><span data-stu-id="5c9c3-108">Errors returned by the call of `displayDialogAsync` because the dialog box cannot be created.</span></span>
- <span data-ttu-id="5c9c3-109">Erros e outros eventos na caixa de diálogo.</span><span class="sxs-lookup"><span data-stu-id="5c9c3-109">Errors, and other events, in the dialog box.</span></span>

## <a name="errors-from-displaydialogasync"></a><span data-ttu-id="5c9c3-110">Erros de displayDialogAsync</span><span class="sxs-lookup"><span data-stu-id="5c9c3-110">Errors from displayDialogAsync</span></span>

<span data-ttu-id="5c9c3-111">Além dos erros gerais da plataforma e do sistema, quatro erros são específicos para chamar `displayDialogAsync` .</span><span class="sxs-lookup"><span data-stu-id="5c9c3-111">In addition to general platform and system errors, four errors are specific to calling `displayDialogAsync`.</span></span>

|<span data-ttu-id="5c9c3-112">Número do código</span><span class="sxs-lookup"><span data-stu-id="5c9c3-112">Code number</span></span>|<span data-ttu-id="5c9c3-113">Significado</span><span class="sxs-lookup"><span data-stu-id="5c9c3-113">Meaning</span></span>|
|:-----|:-----|
|<span data-ttu-id="5c9c3-114">12004</span><span class="sxs-lookup"><span data-stu-id="5c9c3-114">12004</span></span>|<span data-ttu-id="5c9c3-p101">O domínio que a URL transmitiu para `displayDialogAsync` não é confiável. O domínio deve ser o mesmo domínio que o da página de host (incluindo o protocolo e o número de porta).</span><span class="sxs-lookup"><span data-stu-id="5c9c3-p101">The domain of the URL passed to `displayDialogAsync` is not trusted. The domain must be the same domain as the host page (including protocol and port number).</span></span>|
|<span data-ttu-id="5c9c3-117">12005</span><span class="sxs-lookup"><span data-stu-id="5c9c3-117">12005</span></span>|<span data-ttu-id="5c9c3-118">A URL passada para `displayDialogAsync` usa o protocolo HTTP.</span><span class="sxs-lookup"><span data-stu-id="5c9c3-118">The URL passed to `displayDialogAsync` uses the HTTP protocol.</span></span> <span data-ttu-id="5c9c3-119">HTTPS é necessário.</span><span class="sxs-lookup"><span data-stu-id="5c9c3-119">HTTPS is required.</span></span> <span data-ttu-id="5c9c3-120">(Em algumas versões do Office, o texto da mensagem de erro retornado com 12005 é o mesmo retornado para 12004.)</span><span class="sxs-lookup"><span data-stu-id="5c9c3-120">(In some versions of Office, the error message text returned with 12005 is the same one returned for 12004.)</span></span>|
|<span data-ttu-id="5c9c3-121"><span id="12007">12007</span></span><span class="sxs-lookup"><span data-stu-id="5c9c3-121"><span id="12007">12007</span></span></span><!-- The span is needed because office-js-helpers has an error message that links to this table row. -->|<span data-ttu-id="5c9c3-p103">Uma caixa de diálogo já está aberta na janela do host. Uma janela do host, como um painel de tarefas, só pode ter uma caixa de diálogo aberta por vez.</span><span class="sxs-lookup"><span data-stu-id="5c9c3-p103">A dialog box is already opened from this host window. A host window, such as a task pane, can only have one dialog box open at a time.</span></span>|
|<span data-ttu-id="5c9c3-124">12009</span><span class="sxs-lookup"><span data-stu-id="5c9c3-124">12009</span></span>|<span data-ttu-id="5c9c3-125">O usuário opta por ignorar a caixa de diálogo.</span><span class="sxs-lookup"><span data-stu-id="5c9c3-125">The user chose to ignore the dialog box.</span></span> <span data-ttu-id="5c9c3-126">Esse erro pode ocorrer Office na Web, onde os usuários podem optar por não permitir que um complemento apresente uma caixa de diálogo.</span><span class="sxs-lookup"><span data-stu-id="5c9c3-126">This error can occur in Office on the web, where users may choose not to allow an add-in to present a dialog box.</span></span> <span data-ttu-id="5c9c3-127">Para obter mais informações, consulte [Tratamento de bloqueadores pop-up com Office na Web](dialog-best-practices.md#handling-pop-up-blockers-with-office-on-the-web).</span><span class="sxs-lookup"><span data-stu-id="5c9c3-127">For more information, see [Handling pop-up blockers with Office on the web](dialog-best-practices.md#handling-pop-up-blockers-with-office-on-the-web).</span></span>|

<span data-ttu-id="5c9c3-128">Quando `displayDialogAsync` é chamado, ele passa um [objeto AsyncResult](/javascript/api/office/office.asyncresult) para sua função de retorno de chamada.</span><span class="sxs-lookup"><span data-stu-id="5c9c3-128">When `displayDialogAsync` is called, it passes an [AsyncResult](/javascript/api/office/office.asyncresult) object to its callback function.</span></span> <span data-ttu-id="5c9c3-129">Quando a chamada é bem-sucedida, a caixa de diálogo é aberta e `value` a propriedade do objeto é um objeto `AsyncResult` [Dialog.](/javascript/api/office/office.dialog)</span><span class="sxs-lookup"><span data-stu-id="5c9c3-129">When the call is successful, the dialog box is opened, and the `value` property of the `AsyncResult` object is a [Dialog](/javascript/api/office/office.dialog) object.</span></span> <span data-ttu-id="5c9c3-130">Para obter um exemplo disso, consulte [Enviar informações da caixa de diálogo para a página host](dialog-api-in-office-add-ins.md#send-information-from-the-dialog-box-to-the-host-page).</span><span class="sxs-lookup"><span data-stu-id="5c9c3-130">For an example of this, see [Send information from the dialog box to the host page](dialog-api-in-office-add-ins.md#send-information-from-the-dialog-box-to-the-host-page).</span></span> <span data-ttu-id="5c9c3-131">Quando a chamada falha, a caixa de diálogo não é criada, a propriedade do objeto é definida como , e a propriedade do `displayDialogAsync` `status` objeto é `AsyncResult` `Office.AsyncResultStatus.Failed` `error` preenchida.</span><span class="sxs-lookup"><span data-stu-id="5c9c3-131">When the call to `displayDialogAsync` fails, the dialog box is not created, the `status` property of the `AsyncResult` object is set to `Office.AsyncResultStatus.Failed`, and the `error` property of the object is populated.</span></span> <span data-ttu-id="5c9c3-132">Você sempre deve fornecer um retorno de chamada que testa e `status` responde quando é um erro.</span><span class="sxs-lookup"><span data-stu-id="5c9c3-132">You should always provide a callback that tests the `status` and responds when it's an error.</span></span> <span data-ttu-id="5c9c3-133">Para um exemplo que relata a mensagem de erro independentemente do número de código, consulte o código a seguir.</span><span class="sxs-lookup"><span data-stu-id="5c9c3-133">For an example that reports the error message regardless of its code number, see the following code.</span></span> <span data-ttu-id="5c9c3-134">(A `showNotification` função, não definida neste artigo, exibe ou registra o erro.</span><span class="sxs-lookup"><span data-stu-id="5c9c3-134">(The `showNotification` function, not defined in this article, either displays or logs the error.</span></span> <span data-ttu-id="5c9c3-135">Para ver um exemplo de como você pode implementar essa função no seu complemento, consulte Office Exemplo da API de Diálogo de [Complementos](https://github.com/OfficeDev/Office-Add-in-Dialog-API-Simple-Example).)</span><span class="sxs-lookup"><span data-stu-id="5c9c3-135">For an example of how you can implement this function within your add-in, see [Office Add-in Dialog API Example](https://github.com/OfficeDev/Office-Add-in-Dialog-API-Simple-Example).)</span></span>

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

## <a name="errors-and-events-in-the-dialog-box"></a><span data-ttu-id="5c9c3-136">Erros e eventos na caixa de diálogo</span><span class="sxs-lookup"><span data-stu-id="5c9c3-136">Errors and events in the dialog box</span></span>

<span data-ttu-id="5c9c3-137">Três erros e eventos na caixa de diálogo levantarão um `DialogEventReceived` evento na página host.</span><span class="sxs-lookup"><span data-stu-id="5c9c3-137">Three errors and events in the dialog box will raise a `DialogEventReceived` event in the host page.</span></span> <span data-ttu-id="5c9c3-138">Para um lembrete do que é uma página host, consulte Abrir uma caixa [de diálogo de uma página host](dialog-api-in-office-add-ins.md#open-a-dialog-box-from-a-host-page).</span><span class="sxs-lookup"><span data-stu-id="5c9c3-138">For a reminder of what a host page is, see [Open a dialog box from a host page](dialog-api-in-office-add-ins.md#open-a-dialog-box-from-a-host-page).</span></span>

|<span data-ttu-id="5c9c3-139">Número do código</span><span class="sxs-lookup"><span data-stu-id="5c9c3-139">Code number</span></span>|<span data-ttu-id="5c9c3-140">Significado</span><span class="sxs-lookup"><span data-stu-id="5c9c3-140">Meaning</span></span>|
|:-----|:-----|
|<span data-ttu-id="5c9c3-141">12002</span><span class="sxs-lookup"><span data-stu-id="5c9c3-141">12002</span></span>|<span data-ttu-id="5c9c3-142">Uma destas opções:</span><span class="sxs-lookup"><span data-stu-id="5c9c3-142">One of the following:</span></span><br> <span data-ttu-id="5c9c3-143">- Não existe uma página na URL transmitida para `displayDialogAsync`.</span><span class="sxs-lookup"><span data-stu-id="5c9c3-143">- No page exists at the URL that was passed to `displayDialogAsync`.</span></span><br> <span data-ttu-id="5c9c3-144">- A página que foi passada para carregada, mas a caixa de diálogo foi então redirecionada para uma página que não pode encontrar ou carregar, ou foi direcionada para uma URL com sintaxe `displayDialogAsync` inválida.</span><span class="sxs-lookup"><span data-stu-id="5c9c3-144">- The page that was passed to `displayDialogAsync` loaded, but the dialog box was then redirected to a page that it cannot find or load, or it has been directed to a URL with invalid syntax.</span></span>|
|<span data-ttu-id="5c9c3-145">12003</span><span class="sxs-lookup"><span data-stu-id="5c9c3-145">12003</span></span>|<span data-ttu-id="5c9c3-p107">A caixa de diálogo foi direcionada para uma URL com o protocolo HTTP. HTTPS é necessário.</span><span class="sxs-lookup"><span data-stu-id="5c9c3-p107">The dialog box was directed to a URL with the HTTP protocol. HTTPS is required.</span></span>|
|<span data-ttu-id="5c9c3-148">12006</span><span class="sxs-lookup"><span data-stu-id="5c9c3-148">12006</span></span>|<span data-ttu-id="5c9c3-149">A caixa de diálogo foi fechada, geralmente porque o usuário escolheu o **botão Fechar** **X**.</span><span class="sxs-lookup"><span data-stu-id="5c9c3-149">The dialog box was closed, usually because the user chose the **Close** button **X**.</span></span>|

<span data-ttu-id="5c9c3-p108">Seu código pode atribuir um manipulador para o evento `DialogEventReceived` na chamada para `displayDialogAsync`. Apresentamos um exemplo simples a seguir.</span><span class="sxs-lookup"><span data-stu-id="5c9c3-p108">Your code can assign a handler for the `DialogEventReceived` event in the call to `displayDialogAsync`. The following is a simple example.</span></span>

```js
var dialog;
Office.context.ui.displayDialogAsync('https://myDomain/myDialog.html',
    function (result) {
        dialog = result.value;
        dialog.addEventHandler(Office.EventType.DialogEventReceived, processDialogEvent);
    }
);
```

<span data-ttu-id="5c9c3-152">Para um exemplo de um manipulador para o evento que cria mensagens de erro personalizadas para cada código de `DialogEventReceived` erro, consulte o exemplo a seguir.</span><span class="sxs-lookup"><span data-stu-id="5c9c3-152">For an example of a handler for the `DialogEventReceived` event that creates custom error messages for each error code, see the following example.</span></span>

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

<span data-ttu-id="5c9c3-153">Para ver um suplemento de exemplo que manipula erros dessa forma, confira [Exemplo de API de Caixa de diálogo do Suplemento do Office](https://github.com/OfficeDev/Office-Add-in-Dialog-API-Simple-Example).</span><span class="sxs-lookup"><span data-stu-id="5c9c3-153">For a sample add-in that handles errors in this way, see [Office Add-in Dialog API Example](https://github.com/OfficeDev/Office-Add-in-Dialog-API-Simple-Example).</span></span>
