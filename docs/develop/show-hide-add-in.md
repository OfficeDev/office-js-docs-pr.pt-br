---
title: Mostrar ou ocultar o painel de tarefas do seu Add-in do Office
description: Saiba como ocultar ou mostrar programaticamente a interface do usuário de um complemento enquanto ele é executado continuamente.
ms.date: 12/28/2020
localization_priority: Normal
ms.openlocfilehash: 20db609a3a6ded5624391f705dab1ad6b8f6e043
ms.sourcegitcommit: 545888b08f57bb1babb05ccfd83b2b3286bdad5c
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 01/08/2021
ms.locfileid: "49789212"
---
# <a name="show-or-hide-the-task-pane-of-your-office-add-in"></a><span data-ttu-id="d87f2-103">Mostrar ou ocultar o painel de tarefas do seu Add-in do Office</span><span class="sxs-lookup"><span data-stu-id="d87f2-103">Show or hide the task pane of your Office Add-in</span></span>

[!include[Shared JavaScript runtime requirements](../includes/shared-runtime-requirements-note.md)]

<span data-ttu-id="d87f2-104">Você pode mostrar o painel de tarefas do seu Complemento do Office chamando a `Office.addin.showAsTaskpane()` função.</span><span class="sxs-lookup"><span data-stu-id="d87f2-104">You can show the task pane of your Office Add-in by calling the `Office.addin.showAsTaskpane()` function.</span></span>

```javascript
function onCurrentQuarter() {
    Office.addin.showAsTaskpane()
    .then(function() {
        // Code that enables task pane UI elements for
        // working with the current quarter.
    });
}
```

<span data-ttu-id="d87f2-105">O código anterior assume um cenário em que há uma planilha do Excel chamada **CurrentQuarterSales**.</span><span class="sxs-lookup"><span data-stu-id="d87f2-105">The previous code assumes a scenario where there is an Excel worksheet named **CurrentQuarterSales**.</span></span> <span data-ttu-id="d87f2-106">O complemento torna o painel de tarefas visível sempre que essa planilha é ativada.</span><span class="sxs-lookup"><span data-stu-id="d87f2-106">The add-in will make the task pane visible whenever this worksheet is activated.</span></span> <span data-ttu-id="d87f2-107">O método `onCurrentQuarter` é um manipulador para o evento [Office.Worksheet.onActivated](/javascript/api/excel/excel.worksheet?view=excel-js-preview&preserve-view=true#onactivated) que foi registrado para a planilha.</span><span class="sxs-lookup"><span data-stu-id="d87f2-107">The method `onCurrentQuarter` is a handler for the [Office.Worksheet.onActivated](/javascript/api/excel/excel.worksheet?view=excel-js-preview&preserve-view=true#onactivated) event which has been registered for the worksheet.</span></span>

<span data-ttu-id="d87f2-108">Você também pode ocultar o painel de tarefas chamando a `Office.addin.hide()` função.</span><span class="sxs-lookup"><span data-stu-id="d87f2-108">You can also hide the task pane by calling the `Office.addin.hide()` function.</span></span>

```javascript
function onCurrentQuarterDeactivated() {
    Office.addin.hide();
}
```

<span data-ttu-id="d87f2-109">O código anterior é um manipulador que está registrado para o [evento Office.Worksheet.onDeactivated.](/javascript/api/excel/excel.worksheet?view=excel-js-preview&preserve-view=true#ondeactivated)</span><span class="sxs-lookup"><span data-stu-id="d87f2-109">The previous code is a handler that is registered for the [Office.Worksheet.onDeactivated](/javascript/api/excel/excel.worksheet?view=excel-js-preview&preserve-view=true#ondeactivated) event.</span></span>

## <a name="additional-details-on-showing-the-task-pane"></a><span data-ttu-id="d87f2-110">Detalhes adicionais sobre como mostrar o painel de tarefas</span><span class="sxs-lookup"><span data-stu-id="d87f2-110">Additional details on showing the task pane</span></span>

<span data-ttu-id="d87f2-111">Ao chamar, o Office exibirá em um painel de tarefas o arquivo atribuído como o valor da ID do recurso `Office.addin.showAsTaskpane()` ( ) do painel de `resid` tarefas.</span><span class="sxs-lookup"><span data-stu-id="d87f2-111">When you call `Office.addin.showAsTaskpane()`, Office will display in a task pane the file that you assigned as the resource ID (`resid`) value of the task pane.</span></span> <span data-ttu-id="d87f2-112">Esse valor pode ser atribuído ou alterado abrindo seu `resid` **arquivomanifest.xml** e localizando dentro `<SourceLocation>` do `<Action xsi:type="ShowTaskpane">` elemento.</span><span class="sxs-lookup"><span data-stu-id="d87f2-112">This `resid` value can be assigned or changed by opening your **manifest.xml** file and locating `<SourceLocation>` inside the `<Action xsi:type="ShowTaskpane">` element.</span></span>
<span data-ttu-id="d87f2-113">(Confira [Configurar o seu Add-in do Office para usar um tempo de execução compartilhado](configure-your-add-in-to-use-a-shared-runtime.md) para obter detalhes adicionais.)</span><span class="sxs-lookup"><span data-stu-id="d87f2-113">(See [Configure your Office Add-in to use a shared runtime](configure-your-add-in-to-use-a-shared-runtime.md) for additional details.)</span></span>

<span data-ttu-id="d87f2-114">Como `Office.addin.showAsTaskpane()` é um método assíncrono, seu código continuará em execução até que a função seja concluída.</span><span class="sxs-lookup"><span data-stu-id="d87f2-114">Since `Office.addin.showAsTaskpane()` is an asynchronous method, your code will continue running until the function is complete.</span></span> <span data-ttu-id="d87f2-115">Aguarde essa conclusão com a `await` palavra-chave ou um `then()` método, dependendo da sintaxe JavaScript que você está usando.</span><span class="sxs-lookup"><span data-stu-id="d87f2-115">Wait for this completion with either the `await` keyword or a `then()` method, depending on which JavaScript syntax you are using.</span></span>

## <a name="configure-your-add-in-to-use-the-shared-runtime"></a><span data-ttu-id="d87f2-116">Configurar seu complemento para usar o tempo de execução compartilhado</span><span class="sxs-lookup"><span data-stu-id="d87f2-116">Configure your add-in to use the shared runtime</span></span>

<span data-ttu-id="d87f2-117">Para usar os `showAsTaskpane()` métodos `hide()` e os métodos, o seu complemento deve usar o tempo de execução compartilhado.</span><span class="sxs-lookup"><span data-stu-id="d87f2-117">To use the `showAsTaskpane()` and `hide()` methods, your add-in must use the shared runtime.</span></span> <span data-ttu-id="d87f2-118">Para saber mais, confira [Configurar o seu Add-in do Office para usar um tempo de execução compartilhado.](configure-your-add-in-to-use-a-shared-runtime.md)</span><span class="sxs-lookup"><span data-stu-id="d87f2-118">For more information, see [Configure your Office Add-in to use a shared runtime](configure-your-add-in-to-use-a-shared-runtime.md).</span></span>

## <a name="preservation-of-state-and-event-listeners"></a><span data-ttu-id="d87f2-119">Preservação de ouvintes de estado e eventos</span><span class="sxs-lookup"><span data-stu-id="d87f2-119">Preservation of state and event listeners</span></span>

<span data-ttu-id="d87f2-120">The `hide()` and `showAsTaskpane()` methods only change the *visibility* of the task pane.</span><span class="sxs-lookup"><span data-stu-id="d87f2-120">The `hide()` and `showAsTaskpane()` methods only change the *visibility* of the task pane.</span></span> <span data-ttu-id="d87f2-121">Eles não descarregam ou recarregam (ou reinicializam seu estado).</span><span class="sxs-lookup"><span data-stu-id="d87f2-121">They do not unload or reload it (or reinitialize its state).</span></span>

<span data-ttu-id="d87f2-122">Considere o seguinte cenário: um painel de tarefas foi projetado com guias.</span><span class="sxs-lookup"><span data-stu-id="d87f2-122">Consider the following scenario: A task pane is designed with tabs.</span></span> <span data-ttu-id="d87f2-123">A **guia** Início é aberta quando o complemento é lançado pela primeira vez.</span><span class="sxs-lookup"><span data-stu-id="d87f2-123">The **Home** tab is open when the add-in is first launched.</span></span> <span data-ttu-id="d87f2-124">Suponha que um usuário abra a guia **Configurações** e, mais tarde, o código no painel de tarefas chama `hide()` em resposta a algum evento.</span><span class="sxs-lookup"><span data-stu-id="d87f2-124">Suppose a user opens the **Settings** tab and, later, code in the task pane calls `hide()` in response to some event.</span></span> <span data-ttu-id="d87f2-125">Ainda mais tarde, o `showAsTaskpane()` código chama em resposta a outro evento.</span><span class="sxs-lookup"><span data-stu-id="d87f2-125">Still later code calls `showAsTaskpane()` in response to another event.</span></span> <span data-ttu-id="d87f2-126">O painel de tarefas reaparecerá e **a guia Configurações** ainda está selecionada.</span><span class="sxs-lookup"><span data-stu-id="d87f2-126">The task pane will reappear, and the **Settings** tab is still selected.</span></span>

![Uma captura de tela do painel de tarefas que tem quatro guias rotuladas como Página Inicial, Configurações, Favoritos e Contas.](../images/TaskpaneWithTabs.png)

<span data-ttu-id="d87f2-128">Além disso, os ouvintes de eventos registrados no painel de tarefas continuam a ser executados mesmo quando o painel de tarefas está oculto.</span><span class="sxs-lookup"><span data-stu-id="d87f2-128">In addition, any event listeners that are registered in the task pane continue to run even when the task pane is hidden.</span></span>

<span data-ttu-id="d87f2-129">Considere o seguinte cenário: O painel de tarefas tem um manipulador registrado para o Excel e eventos `Worksheet.onActivated` para uma planilha chamada `Worksheet.onDeactivated` **Sheet1**.</span><span class="sxs-lookup"><span data-stu-id="d87f2-129">Consider the following scenario: The task pane has a registered handler for the Excel `Worksheet.onActivated` and `Worksheet.onDeactivated` events for a sheet named **Sheet1**.</span></span> <span data-ttu-id="d87f2-130">O manipulador ativado faz com que um ponto verde apareça no painel de tarefas.</span><span class="sxs-lookup"><span data-stu-id="d87f2-130">The activated handler causes a green dot to appear in the task pane.</span></span> <span data-ttu-id="d87f2-131">O manipulador desativado ativa o ponto vermelho (que é seu estado padrão).</span><span class="sxs-lookup"><span data-stu-id="d87f2-131">The deactivated handler turns the dot red (which is its default state).</span></span> <span data-ttu-id="d87f2-132">Suponha então que o código `hide()` chama **quando Sheet1** não está ativado e o ponto é vermelho.</span><span class="sxs-lookup"><span data-stu-id="d87f2-132">Suppose then that code calls `hide()` when **Sheet1** is not activated and the dot is red.</span></span> <span data-ttu-id="d87f2-133">Enquanto o painel de tarefas está oculto, **Sheet1** é ativado.</span><span class="sxs-lookup"><span data-stu-id="d87f2-133">While the task pane is hidden, **Sheet1** is activated.</span></span> <span data-ttu-id="d87f2-134">Chamadas de código `showAsTaskpane()` posteriores em resposta a algum evento.</span><span class="sxs-lookup"><span data-stu-id="d87f2-134">Later code calls `showAsTaskpane()` in response to some event.</span></span> <span data-ttu-id="d87f2-135">Quando o painel de tarefas é aberto, o ponto fica verde porque os ouvintes e manipuladores de eventos foram embora o painel de tarefas tenha sido ocultado.</span><span class="sxs-lookup"><span data-stu-id="d87f2-135">When the task pane opens, the dot is green because the event listeners and handlers ran even though the task pane was hidden.</span></span>

## <a name="handle-the-visibility-changed-event"></a><span data-ttu-id="d87f2-136">Manipular o evento de visibilidade alterada</span><span class="sxs-lookup"><span data-stu-id="d87f2-136">Handle the visibility changed event</span></span>

<span data-ttu-id="d87f2-137">Quando seu código altera a visibilidade do painel de tarefas com `showAsTaskpane()` `hide()` ou, o Office aciona o `VisibilityModeChanged` evento.</span><span class="sxs-lookup"><span data-stu-id="d87f2-137">When your code changes the visibility of the task pane with `showAsTaskpane()` or `hide()`, Office triggers the `VisibilityModeChanged` event.</span></span> <span data-ttu-id="d87f2-138">Pode ser útil para manipular esse evento.</span><span class="sxs-lookup"><span data-stu-id="d87f2-138">It can be useful to handle this event.</span></span> <span data-ttu-id="d87f2-139">Por exemplo, suponha que o painel de tarefas exibe uma lista de todas as planilhas em uma planilha.</span><span class="sxs-lookup"><span data-stu-id="d87f2-139">For example, suppose the task pane displays a list of all the sheets in a workbook.</span></span> <span data-ttu-id="d87f2-140">Se uma nova planilha for adicionada enquanto o painel de tarefas estiver oculto, tornar o painel de tarefas visível não adicionaria, por si só, o novo nome da planilha à lista.</span><span class="sxs-lookup"><span data-stu-id="d87f2-140">If a new worksheet is added while the task pane is hidden, making the task pane visible would not, in itself, add the new worksheet name to the list.</span></span> <span data-ttu-id="d87f2-141">Mas seu código pode responder ao evento para recarregar a propriedade Worksheet.name de todas as planilhas na coleção `VisibilityModeChanged` [Workbook.worksheets,](/javascript/api/excel/excel.workbook#worksheets) conforme mostrado no código de exemplo abaixo. [](/javascript/api/excel/excel.worksheet#name)</span><span class="sxs-lookup"><span data-stu-id="d87f2-141">But your code can respond to the `VisibilityModeChanged` event to reload the [Worksheet.name](/javascript/api/excel/excel.worksheet#name) property of all the worksheets in the [Workbook.worksheets](/javascript/api/excel/excel.workbook#worksheets) collection as shown in the example code below.</span></span>

<span data-ttu-id="d87f2-142">Para registrar um manipulador para o evento, você não usa um método "adicionar manipulador" como faria na maioria dos contextos javaScript do Office.</span><span class="sxs-lookup"><span data-stu-id="d87f2-142">To register a handler for the event, you do not use an "add handler" method as you would in most Office JavaScript contexts.</span></span> <span data-ttu-id="d87f2-143">Em vez disso, há uma função especial para a qual você passa seu manipulador: [Office.addin.onVisibilityModeChanged](/javascript/api/office/office.addin#onvisibilitymodechanged-listener-).</span><span class="sxs-lookup"><span data-stu-id="d87f2-143">Instead, there is a special function to which you pass your handler: [Office.addin.onVisibilityModeChanged](/javascript/api/office/office.addin#onvisibilitymodechanged-listener-).</span></span> <span data-ttu-id="d87f2-144">Apresentamos um exemplo a seguir.</span><span class="sxs-lookup"><span data-stu-id="d87f2-144">The following is an example.</span></span> <span data-ttu-id="d87f2-145">Observe que a `args.visibilityMode` propriedade é do tipo [VisibilityMode](/javascript/api/office/office.visibilitymode).</span><span class="sxs-lookup"><span data-stu-id="d87f2-145">Note that the `args.visibilityMode` property is type [VisibilityMode](/javascript/api/office/office.visibilitymode).</span></span>

```javascript
Office.addin.onVisibilityModeChanged(function(args) {
    if (args.visibilityMode = "Taskpane"); {
        // Code that runs whenever the task pane is made visible.
        // For example, an Excel.run() that loads the names of
        // all worksheets and passes them to the task pane UI.
    }
});
```

<span data-ttu-id="d87f2-146">A função retorna outra função que *desregula o* manipulador.</span><span class="sxs-lookup"><span data-stu-id="d87f2-146">The function returns another function that *deregisters* the handler.</span></span> <span data-ttu-id="d87f2-147">Veja um exemplo simples, mas não robusto:</span><span class="sxs-lookup"><span data-stu-id="d87f2-147">Here is a simple, but not robust, example:</span></span>

```javascript
var removeVisibilityModeHandler =
    Office.addin.onVisibilityModeChanged(function(args) {
        if (args.visibilityMode = "Taskpane"); {
            // Code that runs whenever the task pane is made visible.
        }
    });


// In some later code path, deregister with:
removeVisibilityModeHandler();
```

<span data-ttu-id="d87f2-148">O método é assíncrono e retorna uma promessa, o que significa que seu código precisa aguardar o cumprimento da promessa antes de chamar o manipulador `onVisibilityModeChanged` **de registro.**</span><span class="sxs-lookup"><span data-stu-id="d87f2-148">The `onVisibilityModeChanged` method is asynchronous and returns a promise, which means that your code needs to await the fulfillment of the promise before it can call the **deregister** handler.</span></span>

```javascript
// await the promise from onVisibilityModeChanged and assign
// the returned deregister handler to removeVisibilityModeHandler.
var removeVisibilityModeHandler =
    await Office.addin.onVisibilityModeChanged(function(args) {
        if (args.visibilityMode = "Taskpane"); {
            // Code that runs whenever the task pane is made visible.
        }
    });
```

<span data-ttu-id="d87f2-149">A função de desregister também é assíncrona e retorna uma promessa.</span><span class="sxs-lookup"><span data-stu-id="d87f2-149">The deregister function is also asynchronous and returns a promise.</span></span> <span data-ttu-id="d87f2-150">Portanto, se você tiver um código que não deve ser executado até que o desregistramento seja concluído, aguarde a promessa retornada pela função de desregister.</span><span class="sxs-lookup"><span data-stu-id="d87f2-150">So, if you have code that should not run until after the deregistration is complete, then you should await the promise returned by the deregister function.</span></span>

```javascript
// await the promise from the deregister handler before continuing
await removeVisibilityModeHandler();
// subsequent code here
```

## <a name="see-also"></a><span data-ttu-id="d87f2-151">Confira também</span><span class="sxs-lookup"><span data-stu-id="d87f2-151">See also</span></span>

- [<span data-ttu-id="d87f2-152">Configurar o Seu Add-in do Office para usar um tempo de execução JavaScript compartilhado</span><span class="sxs-lookup"><span data-stu-id="d87f2-152">Configure your Office Add-in to use a shared JavaScript runtime</span></span>](configure-your-add-in-to-use-a-shared-runtime.md)
- [<span data-ttu-id="d87f2-153">Executar código no seu Add-in do Office quando o documento for aberto</span><span class="sxs-lookup"><span data-stu-id="d87f2-153">Run code in your Office Add-in when the document opens</span></span>](run-code-on-document-open.md)
