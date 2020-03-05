---
title: Mostrar ou ocultar um suplemento do Office em um tempo de execução compartilhado
description: Saiba como ocultar ou mostrar programaticamente a interface do usuário de um suplemento enquanto ele é executado continuamente
ms.date: 03/02/2020
localization_priority: Normal
ms.openlocfilehash: c028823be165723cad3c0b314b53fe7e618188b2
ms.sourcegitcommit: 6c7c98f085dd20f827e0c388e672993412944851
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 03/04/2020
ms.locfileid: "42413788"
---
# <a name="show-or-hide-an-office-add-in-in-a-shared-runtime-preview"></a><span data-ttu-id="ce8aa-103">Mostrar ou ocultar um suplemento do Office em um tempo de execução compartilhado (visualização)</span><span class="sxs-lookup"><span data-stu-id="ce8aa-103">Show or hide an Office Add-in in a shared runtime (preview)</span></span>

<span data-ttu-id="ce8aa-104">Um suplemento do Office pode incluir qualquer uma das seguintes partes:</span><span class="sxs-lookup"><span data-stu-id="ce8aa-104">An Office Add-in can include any of the following parts:</span></span>

- <span data-ttu-id="ce8aa-105">Um painel de tarefas</span><span class="sxs-lookup"><span data-stu-id="ce8aa-105">A task pane</span></span>
- <span data-ttu-id="ce8aa-106">Um arquivo de função sem interface do usuário</span><span class="sxs-lookup"><span data-stu-id="ce8aa-106">A UI-less function file</span></span>
- <span data-ttu-id="ce8aa-107">Uma função personalizada do Excel</span><span class="sxs-lookup"><span data-stu-id="ce8aa-107">An Excel custom function</span></span>

<span data-ttu-id="ce8aa-108">Por padrão, cada parte é executada em seu próprio tempo de execução de JavaScript separado, com seu próprio objeto global e variáveis globais.</span><span class="sxs-lookup"><span data-stu-id="ce8aa-108">By default, each part runs in its own separate JavaScript runtime, with its own global object and global variables.</span></span> 

<span data-ttu-id="ce8aa-109">É possível para suplementos com duas ou mais partes para compartilhar um tempo de execução JavaScript comum.</span><span class="sxs-lookup"><span data-stu-id="ce8aa-109">It's possible for add-ins with two or more parts to share a common JavaScript runtime.</span></span> <span data-ttu-id="ce8aa-110">Esse recurso de tempo de execução compartilhado permite novas APIs de visualização que ocultam e reabrem o painel de tarefas enquanto o suplemento é executado.</span><span class="sxs-lookup"><span data-stu-id="ce8aa-110">This shared runtime feature enables new preview APIs that hide and reopen the task pane while the add-in runs.</span></span>

> [!INCLUDE [Information about using preview APIs](../includes/excel-shared-runtime-preview-note.md)]

## <a name="configure-an-add-in-to-use-a-shared-runtime"></a><span data-ttu-id="ce8aa-111">Configurar um suplemento para usar um tempo de execução compartilhado</span><span class="sxs-lookup"><span data-stu-id="ce8aa-111">Configure an add-in to use a shared runtime</span></span>

<span data-ttu-id="ce8aa-112">Para configurar o suplemento para usar um tempo de execução compartilhado, confira [Configurar o suplemento do Office para usar um tempo de execução compartilhado](configure-your-add-in-to-use-a-shared-runtime.md).</span><span class="sxs-lookup"><span data-stu-id="ce8aa-112">To configure the add-in to use a shared runtime, see [Configure your Office Add-in to use a shared runtime](configure-your-add-in-to-use-a-shared-runtime.md).</span></span>

## <a name="show-and-hide-the-task-pane"></a><span data-ttu-id="ce8aa-113">Mostrar e ocultar o painel de tarefas</span><span class="sxs-lookup"><span data-stu-id="ce8aa-113">Show and hide the task pane</span></span>

<span data-ttu-id="ce8aa-114">As novas APIs estão na `Office.addin` propriedade.</span><span class="sxs-lookup"><span data-stu-id="ce8aa-114">The new APIs are in the `Office.addin` property.</span></span> <span data-ttu-id="ce8aa-115">Para mostrar o painel de tarefas, seu código `Office.addin.showAsTaskpane()`chama.</span><span class="sxs-lookup"><span data-stu-id="ce8aa-115">To show the task pane, your code calls `Office.addin.showAsTaskpane()`.</span></span> <span data-ttu-id="ce8aa-116">O Office será exibido em um painel de tarefas a página que você atribuiu à ID`resid`de recurso () para o painel de tarefas.</span><span class="sxs-lookup"><span data-stu-id="ce8aa-116">Office will display in a task pane the page that you assigned to the resource ID (`resid`) for the task pane.</span></span> <span data-ttu-id="ce8aa-117">Este é o `resid` que você atribuiu ao `<SourceLocation>` do `<Action xsi:type="ShowTaskpane">` manifesto.</span><span class="sxs-lookup"><span data-stu-id="ce8aa-117">This is the `resid` that you assigned to the `<SourceLocation>` of the `<Action xsi:type="ShowTaskpane">` in the manifest.</span></span> <span data-ttu-id="ce8aa-118">(Confira [Configurar o suplemento do Office para usar um tempo de execução compartilhado](configure-your-add-in-to-use-a-shared-runtime.md).)</span><span class="sxs-lookup"><span data-stu-id="ce8aa-118">(See [Configure your Office Add-in to use a shared runtime](configure-your-add-in-to-use-a-shared-runtime.md).)</span></span>

<span data-ttu-id="ce8aa-119">Este é um método assíncrono, portanto, seu código deve aguardar quando o código subsequente não deve ser executado até que seja concluído.</span><span class="sxs-lookup"><span data-stu-id="ce8aa-119">This is an asynchronous method, so your code should await it when the subsequent code should not run until it completes.</span></span> <span data-ttu-id="ce8aa-120">Aguarde essa conclusão com a `await` palavra-chave ou um `then()` método, dependendo da sintaxe JavaScript que você está usando.</span><span class="sxs-lookup"><span data-stu-id="ce8aa-120">Wait for this completion with either the `await` keyword or a `then()` method, depending on which JavaScript syntax you are using.</span></span> <span data-ttu-id="ce8aa-121">O seguinte pressupõe que haja uma planilha do Excel chamada **CurrentQuarterSales**.</span><span class="sxs-lookup"><span data-stu-id="ce8aa-121">The following assumes that there is an Excel worksheet named **CurrentQuarterSales**.</span></span> <span data-ttu-id="ce8aa-122">O suplemento deve tornar o painel de tarefas visível sempre que esta planilha for ativada.</span><span class="sxs-lookup"><span data-stu-id="ce8aa-122">The add-in should make the task pane visible whenever this worksheet is activated.</span></span> <span data-ttu-id="ce8aa-123">O método `onCurrentQuarter` é um manipulador para o evento [Office. Worksheet. OnActivated](/javascript/api/excel/excel.worksheet?view=excel-js-preview#onactivated) que foi registrado para a planilha.</span><span class="sxs-lookup"><span data-stu-id="ce8aa-123">The method `onCurrentQuarter` is a handler for the [Office.Worksheet.onActivated](/javascript/api/excel/excel.worksheet?view=excel-js-preview#onactivated) event which has been registered for the worksheet.</span></span>

```javascript
function onCurrentQuarter() {
    Office.addin.showAsTaskpane()
    .then(function() {
        // Code that enables task pane UI elements for
        // working with the current quarter.
    });
}
```

<span data-ttu-id="ce8aa-124">Para ocultar o painel de tarefas, seu código `Office.addin.hide()`chama.</span><span class="sxs-lookup"><span data-stu-id="ce8aa-124">To hide the task pane, your code calls `Office.addin.hide()`.</span></span> <span data-ttu-id="ce8aa-125">O exemplo a seguir é um manipulador registrado para o evento [Office. Worksheet. OnActivated](/javascript/api/excel/excel.worksheet?view=excel-js-preview#ondeactivated) .</span><span class="sxs-lookup"><span data-stu-id="ce8aa-125">The following example is a handler that is registered for the [Office.Worksheet.onDeactivated](/javascript/api/excel/excel.worksheet?view=excel-js-preview#ondeactivated) event.</span></span>

```javascript
function onCurrentQuarterDeactivated() {
    Office.addin.hide();
}
```

### <a name="preservation-of-state-and-event-listeners"></a><span data-ttu-id="ce8aa-126">Preservação de estado e ouvintes de eventos</span><span class="sxs-lookup"><span data-stu-id="ce8aa-126">Preservation of state and event listeners</span></span>

<span data-ttu-id="ce8aa-127">Os `hide()` métodos `showAsTaskpane()` e só alteram a *visibilidade* do painel de tarefas.</span><span class="sxs-lookup"><span data-stu-id="ce8aa-127">The `hide()` and `showAsTaskpane()` methods only change the *visibility* of the task pane.</span></span> <span data-ttu-id="ce8aa-128">Eles não descarrega ou recarregam (ou reinicializam seu estado).</span><span class="sxs-lookup"><span data-stu-id="ce8aa-128">They do not unload or reload it (or reinitialize its state).</span></span>

<span data-ttu-id="ce8aa-129">Considere o seguinte cenário: um painel de tarefas é projetado com guias.</span><span class="sxs-lookup"><span data-stu-id="ce8aa-129">Consider the following scenario: A task pane is designed with tabs.</span></span> <span data-ttu-id="ce8aa-130">A guia **página inicial** é aberta quando o suplemento é iniciado pela primeira vez.</span><span class="sxs-lookup"><span data-stu-id="ce8aa-130">The **Home** tab is open when the add-in is first launched.</span></span> <span data-ttu-id="ce8aa-131">Suponha que um usuário abra a guia **configurações** e, posteriormente, o código no painel de `hide()` tarefas é chamado em resposta a algum evento.</span><span class="sxs-lookup"><span data-stu-id="ce8aa-131">Suppose a user opens the **Settings** tab and, later, code in the task pane calls `hide()` in response to some event.</span></span> <span data-ttu-id="ce8aa-132">Ainda mais tarde as `showAsTaskpane()` chamadas de código em resposta a outro evento.</span><span class="sxs-lookup"><span data-stu-id="ce8aa-132">Still later code calls `showAsTaskpane()` in response to another event.</span></span> <span data-ttu-id="ce8aa-133">O painel de tarefas será exibido novamente, e a guia **configurações** ainda estará selecionada.</span><span class="sxs-lookup"><span data-stu-id="ce8aa-133">The task pane will reappear, and the **Settings** tab is still selected.</span></span>

![Uma captura de tela do painel de tarefas que tem quatro guias rotuladas Home, configurações, favoritos e contas.](../images/TaskpaneWithTabs.png)

<span data-ttu-id="ce8aa-135">Além disso, todos os ouvintes de eventos registrados no painel de tarefas continuam a ser executados, mesmo quando o painel de tarefas está oculto.</span><span class="sxs-lookup"><span data-stu-id="ce8aa-135">In addition, any event listeners that are registered in the task pane continue to run even when the task pane is hidden.</span></span>

<span data-ttu-id="ce8aa-136">Considere o seguinte cenário: o painel de tarefas tem um manipulador registrado para o `Worksheet.onActivated` Excel `Worksheet.onDeactivated` e eventos para uma planilha chamada **Planilha1**.</span><span class="sxs-lookup"><span data-stu-id="ce8aa-136">Consider the following scenario: The task pane has a registered handler for the Excel `Worksheet.onActivated` and `Worksheet.onDeactivated` events for a sheet named **Sheet1**.</span></span> <span data-ttu-id="ce8aa-137">O manipulador ativado faz com que um ponto verde apareça no painel de tarefas.</span><span class="sxs-lookup"><span data-stu-id="ce8aa-137">The activated handler causes a green dot to appear in the task pane.</span></span> <span data-ttu-id="ce8aa-138">O manipulador desativado transforma o ponto vermelho (que é seu estado padrão).</span><span class="sxs-lookup"><span data-stu-id="ce8aa-138">The deactivated handler turns the dot red (which is its default state).</span></span> <span data-ttu-id="ce8aa-139">Suponha que o código chame `hide()` quando **Sheet1** não está ativado e o ponto é vermelho.</span><span class="sxs-lookup"><span data-stu-id="ce8aa-139">Suppose then that code calls `hide()` when **Sheet1** is not activated and the dot is red.</span></span> <span data-ttu-id="ce8aa-140">Enquanto o painel de tarefas está oculto, **Sheet1** é ativada.</span><span class="sxs-lookup"><span data-stu-id="ce8aa-140">While the task pane is hidden, **Sheet1** is activated.</span></span> <span data-ttu-id="ce8aa-141">Chamadas `showAsTaskpane()` de código posteriores em resposta a algum evento.</span><span class="sxs-lookup"><span data-stu-id="ce8aa-141">Later code calls `showAsTaskpane()` in response to some event.</span></span> <span data-ttu-id="ce8aa-142">Quando o painel de tarefas é aberto, o ponto é verde porque os ouvintes e manipuladores de eventos foram executados, mesmo que o painel de tarefas tenha sido ocultado.</span><span class="sxs-lookup"><span data-stu-id="ce8aa-142">When the task pane opens, the dot is green because the event listeners and handlers ran even though the task pane was hidden.</span></span>

### <a name="handle-visibility-changed-event"></a><span data-ttu-id="ce8aa-143">Evento alterado de visibilidade de manipulação</span><span class="sxs-lookup"><span data-stu-id="ce8aa-143">Handle visibility changed event</span></span>

<span data-ttu-id="ce8aa-144">Quando o código altera a visibilidade do painel de tarefas com `showAsTaskpane()` ou `hide()`, o Office aciona `VisibilityModeChanged` o evento.</span><span class="sxs-lookup"><span data-stu-id="ce8aa-144">When your code changes the visibility of the task pane with `showAsTaskpane()` or `hide()`, Office triggers the `VisibilityModeChanged` event.</span></span> <span data-ttu-id="ce8aa-145">Pode ser útil para lidar com esse evento.</span><span class="sxs-lookup"><span data-stu-id="ce8aa-145">It can be useful to handle this event.</span></span> <span data-ttu-id="ce8aa-146">Por exemplo, suponha que o painel de tarefas exiba uma lista de todas as planilhas em uma pasta de trabalho.</span><span class="sxs-lookup"><span data-stu-id="ce8aa-146">For example, suppose the task pane displays a list of all the sheets in a workbook.</span></span> <span data-ttu-id="ce8aa-147">Se uma nova planilha for adicionada enquanto o painel de tarefas estiver oculto, tornar o painel de tarefas visível não, sozinho, adicionará o novo nome da planilha à lista.</span><span class="sxs-lookup"><span data-stu-id="ce8aa-147">If a new worksheet is added while the task pane is hidden, making the task pane visible would not, in itself, add the new worksheet name to the list.</span></span> <span data-ttu-id="ce8aa-148">Mas seu código pode responder ao `VisibilityModeChanged` evento para recarregar a propriedade [Worksheet.Name](/javascript/api/excel/excel.worksheet#name) de todas as planilhas na coleção [Workbook. Worksheets](/javascript/api/excel/excel.workbook#worksheets) , conforme mostrado no exemplo de código abaixo.</span><span class="sxs-lookup"><span data-stu-id="ce8aa-148">But your code can respond to the `VisibilityModeChanged` event to reload the [Worksheet.name](/javascript/api/excel/excel.worksheet#name) property of all the worksheets in the [Workbook.worksheets](/javascript/api/excel/excel.workbook#worksheets) collection as shown in the example code below.</span></span>

<span data-ttu-id="ce8aa-149">Para registrar um manipulador para o evento, você não usa um método "Add Handler" como faria na maioria dos contextos JavaScript do Office.</span><span class="sxs-lookup"><span data-stu-id="ce8aa-149">To register a handler for the event, you do not use an "add handler" method as you would in most Office JavaScript contexts.</span></span> <span data-ttu-id="ce8aa-150">Em vez disso, há uma função especial para a qual você passa seu manipulador: [Office. AddIn. onVisibilityModeChanged](/javascript/api/office/office.addin#onvisibilitymodechanged-listener-).</span><span class="sxs-lookup"><span data-stu-id="ce8aa-150">Instead, there is a special function to which you pass your handler: [Office.addin.onVisibilityModeChanged](/javascript/api/office/office.addin#onvisibilitymodechanged-listener-).</span></span> <span data-ttu-id="ce8aa-151">Apresentamos um exemplo a seguir.</span><span class="sxs-lookup"><span data-stu-id="ce8aa-151">The following is an example.</span></span> <span data-ttu-id="ce8aa-152">Observe que a `args.visibilityMode` propriedade é do tipo [VisibilityMode](/javascript/api/office/office.visibilitymode).</span><span class="sxs-lookup"><span data-stu-id="ce8aa-152">Note that the `args.visibilityMode` property is type [VisibilityMode](/javascript/api/office/office.visibilitymode).</span></span>

```javascript
Office.addin.onVisibilityModeChanged(function(args) {
    if (args.visibilityMode = "Taskpane"); {
        // Code that runs whenever the task pane is made visible.
        // For example, an Excel.run() that loads the names of
        // all worksheets and passes them to the task pane UI.
    }
});
```

<span data-ttu-id="ce8aa-153">A função retorna outra função que *cancela o registro* do manipulador.</span><span class="sxs-lookup"><span data-stu-id="ce8aa-153">The function returns another function that *deregisters* the handler.</span></span> <span data-ttu-id="ce8aa-154">Veja um exemplo simples, mas não robusto:</span><span class="sxs-lookup"><span data-stu-id="ce8aa-154">Here is a simple, but not robust, example:</span></span>

```javascript
var removeVisibilityModeHandler =
    Office.addin.onVisibilityModeChanged(function(args) {
        if (args.visibilityMode = "Taskpane"); {
            // Code that runs whenever the task pane is made visible.
        }
    });


// In some later code path, deregister with:
removeVisibilityModeHandler();
```

<span data-ttu-id="ce8aa-155">O `onVisibilityModeChanged` método é assíncrono, o que significa que, se seu código chama o manipulador `onVisibilityModeChanged` de *cancelamento de registro* que retorna, você deve garantir que `onVisibilityModeChanged` tenha sido concluída antes de chamar o manipulador de cancelamento de registro.</span><span class="sxs-lookup"><span data-stu-id="ce8aa-155">The `onVisibilityModeChanged` method is asynchronous which means that if your code calls the *deregister* handler that `onVisibilityModeChanged` returns, you should ensure that `onVisibilityModeChanged` has completed before calling the deregister handler.</span></span> <span data-ttu-id="ce8aa-156">Uma maneira de fazer isso é usar a `await` palavra-chave no método Call como no exemplo a seguir.</span><span class="sxs-lookup"><span data-stu-id="ce8aa-156">One way to do that is to use the `await` keyword on the method call as in the following example.</span></span>

```javascript
var removeVisibilityModeHandler =
    await Office.addin.onVisibilityModeChanged(function(args) {
        if (args.visibilityMode = "Taskpane"); {
            // Code that runs whenever the task pane is made visible.
        }
    });
```

<span data-ttu-id="ce8aa-157">Se você quiser usar apenas JavaScript ES2015, seu código poderá usar o `then` método para aguardar até que o objeto Promise retornado tenha sido resolvido e atribuir a função retornada a uma variável global, como no exemplo a seguir.</span><span class="sxs-lookup"><span data-stu-id="ce8aa-157">If you want to use only pre-ES2015 JavaScript, your code can use the `then` method to wait until the returned Promise object has resolved and assign the returned function to a global variable as in the following example.</span></span>

```javascript
var removeVisibilityModeHandler;

Office.addin.onVisibilityModeChanged(function(args) {
        if (args.visibilityMode = "Taskpane"); {
            // Code that runs whenever the task pane is made visible.
        }
}).then(function(removeHandler) {
        removeVisibilityModeHandler = removeHandler;
    });

// In some later code path, deregister with:
removeVisibilityModeHandler();
```

<span data-ttu-id="ce8aa-158">A função cancelamento de registro é assíncrona.</span><span class="sxs-lookup"><span data-stu-id="ce8aa-158">The deregister function is itself asynchronous.</span></span> <span data-ttu-id="ce8aa-159">Portanto, se você tiver um código que não deve ser executado até que o cancelamento de registro seja concluído, a função de cancelamento também deverá `await` ser esperada com `then` a palavra-chave ou com um método como nos exemplos a seguir.</span><span class="sxs-lookup"><span data-stu-id="ce8aa-159">So, if you have code that should not run until after the deregistration is complete, then the deregister function should also be awaited with either the `await` keyword or with a `then` method as in the following examples.</span></span>

<span data-ttu-id="ce8aa-160">Para cancelar o registro do manipulador:</span><span class="sxs-lookup"><span data-stu-id="ce8aa-160">To deregister the handler:</span></span>

```javascript
await removeVisibilityModeHandler();
// subsequent code here

// or use pre-ES2015 syntax:
removeVisibilityModeHandler().then(function () {
        // subsequent code here
    })
```
