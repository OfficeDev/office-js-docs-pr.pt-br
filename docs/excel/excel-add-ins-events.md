---
title: Trabalhar com eventos usando a API JavaScript do Excel
description: ''
ms.date: 10/11/2019
localization_priority: Priority
ms.openlocfilehash: 1838ddf2016d5c0d4651991ce569fd98d6ac960e
ms.sourcegitcommit: 78bbbd6cb5a270164b26038675a222defc3be55e
ms.translationtype: HT
ms.contentlocale: pt-BR
ms.lasthandoff: 10/11/2019
ms.locfileid: "37471350"
---
# <a name="work-with-events-using-the-excel-javascript-api"></a><span data-ttu-id="e7535-102">Trabalhar com eventos usando a API JavaScript do Excel</span><span class="sxs-lookup"><span data-stu-id="e7535-102">Work with Events using the Excel JavaScript API</span></span>

<span data-ttu-id="e7535-103">Este artigo descreve conceitos importantes relacionados ao trabalho com eventos no Excel e fornece exemplos de código que mostram como registrar manipuladores de eventos, lidar com eventos e remover manipuladores de eventos usando as APIs JavaScript do Excel.</span><span class="sxs-lookup"><span data-stu-id="e7535-103">This article describes important concepts related to working with events in Excel and provides code samples that show how to register event handlers, handle events, and remove event handlers using the Excel JavaScript API.</span></span>

## <a name="events-in-excel"></a><span data-ttu-id="e7535-104">Eventos no Excel</span><span class="sxs-lookup"><span data-stu-id="e7535-104">Events in Excel</span></span>

<span data-ttu-id="e7535-p101">Sempre que ocorrerem certos tipos de alterações em uma pasta de trabalho do Excel, uma notificação do evento será ativada. Ao usar as APIs JavaScript do Excel, você pode registrar manipuladores de eventos que permitem que o suplemento execute automaticamente uma função designada quando ocorre um evento específico. Os eventos a seguir têm suporte no momento:</span><span class="sxs-lookup"><span data-stu-id="e7535-p101">Each time certain types of changes occur in an Excel workbook, an event notification fires. By using the Excel JavaScript API, you can register event handlers that allow your add-in to automatically run a designated function when a specific event occurs. The following events are currently supported.</span></span>

| <span data-ttu-id="e7535-108">Evento</span><span class="sxs-lookup"><span data-stu-id="e7535-108">Event</span></span> | <span data-ttu-id="e7535-109">Descrição</span><span class="sxs-lookup"><span data-stu-id="e7535-109">Description</span></span> | <span data-ttu-id="e7535-110">Objetos com suporte</span><span class="sxs-lookup"><span data-stu-id="e7535-110">Supported objects</span></span> |
|:---------------|:-------------|:-----------|
| `onActivated` | <span data-ttu-id="e7535-111">Ocorre quando um objeto está ativado.</span><span class="sxs-lookup"><span data-stu-id="e7535-111">Occurs when an object is activated.</span></span> | <span data-ttu-id="e7535-112">[**Gráfico**](/javascript/api/excel/excel.chart), [**ChartCollection**](/javascript/api/excel/excel.chartcollection), [**Shape**](/javascript/api/excel/excel.shape), [**Planilha**](/javascript/api/excel/excel.worksheet), [**WorksheetCollection**](/javascript/api/excel/excel.worksheetcollection)</span><span class="sxs-lookup"><span data-stu-id="e7535-112">[**Chart**](/javascript/api/excel/excel.chart), [**ChartCollection**](/javascript/api/excel/excel.chartcollection), [**Worksheet**](/javascript/api/excel/excel.shape), [**WorksheetCollection**](/javascript/api/excel/excel.worksheet)</span></span> |
| `onAdded` | <span data-ttu-id="e7535-113">Ocorre quando um objeto é adicionado à coleção.</span><span class="sxs-lookup"><span data-stu-id="e7535-113">Occurs when a view is added to the collection.</span></span> | <span data-ttu-id="e7535-114">[**ChartCollection**](/javascript/api/excel/excel.chartcollection), [**TableCollection**](/javascript/api/excel/excel.tablecollection), [**WorksheetCollection**](/javascript/api/excel/excel.worksheetcollection)</span><span class="sxs-lookup"><span data-stu-id="e7535-114">[**ChartCollection**](/javascript/api/excel/excel.chartcollection), [**TableCollection**](/javascript/api/excel/excel.tablecollection), [**WorksheetCollection**](/javascript/api/excel/excel.worksheetcollection)</span></span> |
| `onAutoSaveSettingChanged` | <span data-ttu-id="e7535-115">Ocorre quando a `autoSave` configuração é alterada na pasta de trabalho.</span><span class="sxs-lookup"><span data-stu-id="e7535-115">Occurs when the `autoSave` setting is changed on the workbook.</span></span> | [<span data-ttu-id="e7535-116">**Workbook**</span><span class="sxs-lookup"><span data-stu-id="e7535-116">**Workbook**</span></span>](/javascript/api/excel/excel.workbook) |
| `onCalculated` | <span data-ttu-id="e7535-117">Ocorre quando uma planilha terminou um cálculo (ou todas as planilhas do conjunto terminaram).</span><span class="sxs-lookup"><span data-stu-id="e7535-117">Occurs when a worksheet has finished calculation (or all the worksheets of the collection have finished).</span></span> | <span data-ttu-id="e7535-118">[**WorksheetCollection**](/javascript/api/excel/excel.worksheet), [**Planilha**](/javascript/api/excel/excel.worksheetcollection)</span><span class="sxs-lookup"><span data-stu-id="e7535-118">[**Worksheet**](/javascript/api/excel/excel.worksheet), [**WorksheetCollection**](/javascript/api/excel/excel.worksheetcollection)</span></span> |
| `onChanged` | <span data-ttu-id="e7535-119">Ocorre quando os dados das células são alterados.</span><span class="sxs-lookup"><span data-stu-id="e7535-119">Occurs when data within cells is changed.</span></span> | <span data-ttu-id="e7535-120">[**Tabela**](/javascript/api/excel/excel.table), [**TableCollection**](/javascript/api/excel/excel.tablecollection), [**Planilha**](/javascript/api/excel/excel.worksheet), [**WorksheetCollection**](/javascript/api/excel/excel.worksheetcollection)</span><span class="sxs-lookup"><span data-stu-id="e7535-120">[**Table**](/javascript/api/excel/excel.table), [**TableCollection**](/javascript/api/excel/excel.tablecollection), [**Worksheet**](/javascript/api/excel/excel.worksheet), [**WorksheetCollection**](/javascript/api/excel/excel.worksheetcollection)</span></span> |
| `onDataChanged` | <span data-ttu-id="e7535-121">Ocorre quando os dados ou a formatação dentro da associação são alterados.</span><span class="sxs-lookup"><span data-stu-id="e7535-121">Occurs when data or formatting within the binding is changed.</span></span> | [<span data-ttu-id="e7535-122">**Associação**</span><span class="sxs-lookup"><span data-stu-id="e7535-122">**Binding**</span></span>](/javascript/api/excel/excel.binding) |
| `onDeactivated` | <span data-ttu-id="e7535-123">Ocorre quando um objeto é desativado.</span><span class="sxs-lookup"><span data-stu-id="e7535-123">Occurs when an object is deactivated.</span></span> | <span data-ttu-id="e7535-124">[**Gráfico**](/javascript/api/excel/excel.chart), [**ChartCollection**](/javascript/api/excel/excel.chartcollection), [**Shape**](/javascript/api/excel/excel.shape), [**Planilha**](/javascript/api/excel/excel.worksheet), [**WorksheetCollection**](/javascript/api/excel/excel.worksheetcollection)</span><span class="sxs-lookup"><span data-stu-id="e7535-124">[**Chart**](/javascript/api/excel/excel.chart), [**ChartCollection**](/javascript/api/excel/excel.chartcollection), [**Worksheet**](/javascript/api/excel/excel.shape), [**WorksheetCollection**](/javascript/api/excel/excel.worksheet)</span></span> |
| `onDeleted` | <span data-ttu-id="e7535-125">Ocorre quando um objeto é excluído da coleção.</span><span class="sxs-lookup"><span data-stu-id="e7535-125">Occurs when an item is deleted from the specified collection.</span></span> | <span data-ttu-id="e7535-126">[**ChartCollection**](/javascript/api/excel/excel.chartcollection), [**TableCollection**](/javascript/api/excel/excel.tablecollection), [**WorksheetCollection**](/javascript/api/excel/excel.worksheetcollection)</span><span class="sxs-lookup"><span data-stu-id="e7535-126">[**ChartCollection**](/javascript/api/excel/excel.chartcollection), [**TableCollection**](/javascript/api/excel/excel.tablecollection), [**WorksheetCollection**](/javascript/api/excel/excel.worksheetcollection)</span></span> |
| `onFormatChanged` | <span data-ttu-id="e7535-127">Ocorre quando o formato é alterado em uma planilha.</span><span class="sxs-lookup"><span data-stu-id="e7535-127">Occurs when the format is changed on a worksheet.</span></span> | <span data-ttu-id="e7535-128">[**WorksheetCollection**](/javascript/api/excel/excel.worksheet), [**Planilha**](/javascript/api/excel/excel.worksheetcollection)</span><span class="sxs-lookup"><span data-stu-id="e7535-128">[**Worksheet**](/javascript/api/excel/excel.worksheet), [**WorksheetCollection**](/javascript/api/excel/excel.worksheetcollection)</span></span> |
| `onSelectionChanged` | <span data-ttu-id="e7535-129">Ocorre quando uma célula ativa ou um intervalo selecionado são alterados.</span><span class="sxs-lookup"><span data-stu-id="e7535-129">Occurs when the active cell or selected range is changed.</span></span> | <span data-ttu-id="e7535-130">[**Associação**](/javascript/api/excel/excel.binding), [**Tabela**](/javascript/api/excel/excel.table),  [**Worksheet**](/javascript/api/excel/excel.worksheet), [**WorksheetCollection**](/javascript/api/excel/excel.worksheetcollection)</span><span class="sxs-lookup"><span data-stu-id="e7535-130">[**Table**](/javascript/api/excel/excel.binding), [**TableCollection**](/javascript/api/excel/excel.table), [**Worksheet**](/javascript/api/excel/excel.worksheet), [**WorksheetCollection**](/javascript/api/excel/excel.worksheetcollection)</span></span> |
| `onSettingsChanged` | <span data-ttu-id="e7535-131">Ocorre quando as Configurações no documento são alteradas.</span><span class="sxs-lookup"><span data-stu-id="e7535-131">Occurs when the Settings in the document are changed.</span></span> | [<span data-ttu-id="e7535-132">**SettingCollection**</span><span class="sxs-lookup"><span data-stu-id="e7535-132">**SettingCollection**</span></span>](/javascript/api/excel/excel.settingcollection) |

> [!WARNING]
> <span data-ttu-id="e7535-133">O `onSelectionChanged` atualmente é instável.</span><span class="sxs-lookup"><span data-stu-id="e7535-133">`onSelectionChanged` is currently unstable.</span></span> <span data-ttu-id="e7535-134">Existe uma solução alternativa para o uso confiável de `onSelectionChanged`.</span><span class="sxs-lookup"><span data-stu-id="e7535-134">There is a workaround to reliably use `onSelectionChanged`.</span></span> <span data-ttu-id="e7535-135">Adicione o seguinte código à seção `<head>` da sua home page HTML:</span><span class="sxs-lookup"><span data-stu-id="e7535-135">Add the following code to the `<head>` section of your HTML home page:</span></span>
>
> ```HTML
> <script> MutationObserver=null; </script>
> ```
>
> <span data-ttu-id="e7535-136">Uma discussão completa sobre o assunto pode ser encontrada no [repositório office-js GitHub](https://github.com/OfficeDev/office-js/issues/533).</span><span class="sxs-lookup"><span data-stu-id="e7535-136">A full discussion of the issue can be found on the [office-js GitHub repo](https://github.com/OfficeDev/office-js/issues/533).</span></span>

### <a name="events-in-preview"></a><span data-ttu-id="e7535-137">Eventos no modo de visualização</span><span class="sxs-lookup"><span data-stu-id="e7535-137">Events in preview</span></span>

> [!NOTE]
> <span data-ttu-id="e7535-138">Os seguintes eventos estão disponíveis atualmente apenas na visualização pública.</span><span class="sxs-lookup"><span data-stu-id="e7535-138">The following events are currently available only in public preview.</span></span> [!INCLUDE [Information about using preview APIs](../includes/using-excel-preview-apis.md)]

| <span data-ttu-id="e7535-139">Evento</span><span class="sxs-lookup"><span data-stu-id="e7535-139">Event</span></span> | <span data-ttu-id="e7535-140">Descrição</span><span class="sxs-lookup"><span data-stu-id="e7535-140">Description</span></span> | <span data-ttu-id="e7535-141">Objetos com suporte</span><span class="sxs-lookup"><span data-stu-id="e7535-141">Supported objects</span></span> |
|:---------------|:-------------|:-----------|
| `onColumnSorted` | <span data-ttu-id="e7535-142">Ocorre quando uma ou mais colunas são classificadas.</span><span class="sxs-lookup"><span data-stu-id="e7535-142">Occurs when one or more columns have been sorted.</span></span> <span data-ttu-id="e7535-143">Isso acontece como resultado de uma operação de classificação da esquerda para a direita.</span><span class="sxs-lookup"><span data-stu-id="e7535-143">This happens as the result of a left-to-right sort operation.</span></span> | <span data-ttu-id="e7535-144">[**Planilha**](/javascript/api/excel/excel.worksheet), [**WorksheetCollection**](/javascript/api/excel/excel.worksheetcollection)</span><span class="sxs-lookup"><span data-stu-id="e7535-144">[**Worksheet**](/javascript/api/excel/excel.worksheet), [**WorksheetCollection**](/javascript/api/excel/excel.worksheetcollection)</span></span> |
| `onFiltered` | <span data-ttu-id="e7535-145">Ocorre quando um filtro é aplicado a um objeto.</span><span class="sxs-lookup"><span data-stu-id="e7535-145">Occurs when filter is applied on an object.</span></span> | <span data-ttu-id="e7535-146">[**Tabela**](/javascript/api/excel/excel.table), [**TableCollection**](/javascript/api/excel/excel.tablecollection), [**Planilha**](/javascript/api/excel/excel.worksheet), [**WorksheetCollection**](/javascript/api/excel/excel.worksheetcollection)</span><span class="sxs-lookup"><span data-stu-id="e7535-146">[**Table**](/javascript/api/excel/excel.table), [**TableCollection**](/javascript/api/excel/excel.tablecollection), [**Worksheet**](/javascript/api/excel/excel.worksheet), [**WorksheetCollection**](/javascript/api/excel/excel.worksheetcollection)</span></span> |
| `onRowHiddenChanged` | <span data-ttu-id="e7535-147">Ocorre quando o estado de linha oculta é alterado em uma planilha específica.</span><span class="sxs-lookup"><span data-stu-id="e7535-147">Occurs when the row-hidden state changes on a specific worksheet.</span></span> | <span data-ttu-id="e7535-148">[**Planilha**](/javascript/api/excel/excel.worksheet), [**WorksheetCollection**](/javascript/api/excel/excel.worksheetcollection)</span><span class="sxs-lookup"><span data-stu-id="e7535-148">[**Worksheet**](/javascript/api/excel/excel.worksheet), [**WorksheetCollection**](/javascript/api/excel/excel.worksheetcollection)</span></span> |
| `onRowSorted` | <span data-ttu-id="e7535-149">Ocorre quando uma ou mais linhas são classificadas.</span><span class="sxs-lookup"><span data-stu-id="e7535-149">Occurs when one or more rows have been sorted.</span></span> <span data-ttu-id="e7535-150">Isso ocorre como resultado de uma operação de classificação de cima para baixo.</span><span class="sxs-lookup"><span data-stu-id="e7535-150">This happens as the result of a top-to-bottom sort operation.</span></span> | <span data-ttu-id="e7535-151">[**Planilha**](/javascript/api/excel/excel.worksheet), [**WorksheetCollection**](/javascript/api/excel/excel.worksheetcollection)</span><span class="sxs-lookup"><span data-stu-id="e7535-151">[**Worksheet**](/javascript/api/excel/excel.worksheet), [**WorksheetCollection**](/javascript/api/excel/excel.worksheetcollection)</span></span> |
| `onSingleClicked` | <span data-ttu-id="e7535-152">Acontece quando a operação é clicada/pressionada com o botão esquerdo do mouse ocorre na planilha.</span><span class="sxs-lookup"><span data-stu-id="e7535-152">Occurs when left-clicked/tapped operation happens in the worksheet.</span></span> | <span data-ttu-id="e7535-153">[**Planilha**](/javascript/api/excel/excel.worksheet), [**WorksheetCollection**](/javascript/api/excel/excel.worksheetcollection)</span><span class="sxs-lookup"><span data-stu-id="e7535-153">[**Worksheet**](/javascript/api/excel/excel.worksheet), [**WorksheetCollection**](/javascript/api/excel/excel.worksheetcollection)</span></span> |

### <a name="event-triggers"></a><span data-ttu-id="e7535-154">Gatilhos de eventos</span><span class="sxs-lookup"><span data-stu-id="e7535-154">Event triggers</span></span>

<span data-ttu-id="e7535-155">Os eventos em uma pasta de trabalho do Excel podem ser acionados por:</span><span class="sxs-lookup"><span data-stu-id="e7535-155">Events within an Excel workbook can be triggered by:</span></span>

- <span data-ttu-id="e7535-156">Interação do usuário por meio da interface (IU) do Excel que altera a pasta de trabalho</span><span class="sxs-lookup"><span data-stu-id="e7535-156">User interaction via the Excel user interface (UI) that changes the workbook</span></span>
- <span data-ttu-id="e7535-157">Código de suplemento do Office (JavaScript) que altera a pasta de trabalho</span><span class="sxs-lookup"><span data-stu-id="e7535-157">Office Add-in (JavaScript) code that changes the workbook</span></span>
- <span data-ttu-id="e7535-158">Código de suplemento VBA (macro) que altera a pasta de trabalho</span><span class="sxs-lookup"><span data-stu-id="e7535-158">VBA add-in (macro) code that changes the workbook</span></span>

<span data-ttu-id="e7535-159">Todas as alterações que sejam compatíveis com o comportamento padrão do Excel acionarão eventos correspondentes em uma pasta de trabalho.</span><span class="sxs-lookup"><span data-stu-id="e7535-159">Any change that complies with default behavior of Excel will trigger the corresponding event(s) in a workbook.</span></span>

### <a name="lifecycle-of-an-event-handler"></a><span data-ttu-id="e7535-160">Ciclo de vida de um manipulador de eventos</span><span class="sxs-lookup"><span data-stu-id="e7535-160">Lifecycle of an event handler</span></span>

<span data-ttu-id="e7535-161">Um manipulador de eventos é criado quando um suplemento registra o manipulador de eventos.</span><span class="sxs-lookup"><span data-stu-id="e7535-161">An event handler is created when an add-in registers the event handler.</span></span> <span data-ttu-id="e7535-162">Ele é destruído quando o suplemento cancela o registro de manipulador de evento ou quando o suplemento é atualizado, recarregado ou fechado.</span><span class="sxs-lookup"><span data-stu-id="e7535-162">It is destroyed when the add-in unregisters the event handler or when the add-in is refreshed, reloaded, or closed.</span></span> <span data-ttu-id="e7535-163">Manipuladores de eventos não são mantidos como parte do arquivo do Excel ou em sessões do Excel Online.</span><span class="sxs-lookup"><span data-stu-id="e7535-163">Event handlers do not persist as part of the Excel file, or across sessions with Excel Online.</span></span>

> [!CAUTION]
> <span data-ttu-id="e7535-164">Quando um objeto ao qual os eventos são registrados é excluído (por exemplo, uma tabela com um `onChanged` evento registrado), o manipulador de eventos não disparará mais, mas permanecerá na memória até que o suplemento ou sessão do Excel atualize ou feche.</span><span class="sxs-lookup"><span data-stu-id="e7535-164">When an object to which events are registered is deleted (e.g., a table with an `onChanged` event registered), the event handler no longer triggers but remains in memory until the add-in or Excel session refreshes or closes.</span></span>

### <a name="events-and-coauthoring"></a><span data-ttu-id="e7535-165">Eventos e coautoria</span><span class="sxs-lookup"><span data-stu-id="e7535-165">Events and coauthoring</span></span>

<span data-ttu-id="e7535-p107">Com a [coautoria](co-authoring-in-excel-add-ins.md), várias pessoas podem trabalhar em conjunto e editar a mesma pasta de trabalho do Excel simultaneamente. Em eventos que podem ser disparados por um coautor, como `onChanged`, o objeto de **evento** respectivo conterá a propriedade **fonte** que indica se o evento foi acionado localmente pelo usuário atual (`event.source = Local`) ou pelo coautor remoto (`event.source = Remote`).</span><span class="sxs-lookup"><span data-stu-id="e7535-p107">With [coauthoring](co-authoring-in-excel-add-ins.md), multiple people can work together and edit the same Excel workbook simultaneously. For events that can be triggered by a coauthor, such as `onChanged`, the corresponding **Event** object will contain a **source** property that indicates whether the event was triggered locally by the current user (`event.source = Local`) or was triggered by the remote coauthor (`event.source = Remote`).</span></span>

## <a name="register-an-event-handler"></a><span data-ttu-id="e7535-168">Registrar um manipulador de eventos.</span><span class="sxs-lookup"><span data-stu-id="e7535-168">Register an event handler</span></span>

<span data-ttu-id="e7535-p108">O exemplo de código a seguir registra um manipulador de eventos para o evento `onChanged` na planilha **Sample**. O código especifica que, quando os dados forem alterados na planilha, a função `handleDataChange` deve ser executada.</span><span class="sxs-lookup"><span data-stu-id="e7535-p108">The following code sample registers an event handler for the `onChanged` event in the worksheet named **Sample**. The code specifies that when data changes in that worksheet, the `handleDataChange` function should run.</span></span>

```js
Excel.run(function (context) {
    var worksheet = context.workbook.worksheets.getItem("Sample");
    worksheet.onChanged.add(handleChange);

    return context.sync()
        .then(function () {
            console.log("Event handler successfully registered for onChanged event in the worksheet.");
        });
}).catch(errorHandlerFunction);
```

## <a name="handle-an-event"></a><span data-ttu-id="e7535-171">Manipular um evento</span><span class="sxs-lookup"><span data-stu-id="e7535-171">Handle an event</span></span>

<span data-ttu-id="e7535-p109">Como mostrado no exemplo anterior, quando você registrar um manipulador de eventos, indica a função a ser executada quando o evento especificado ocorre. Você pode criar essa função para executar as ações que seu cenário exige. O exemplo de código a seguir mostra uma função de manipulador de eventos que simplesmente grava informações sobre o evento no console.</span><span class="sxs-lookup"><span data-stu-id="e7535-p109">As shown in the previous example, when you register an event handler, you indicate the function that should run when the specified event occurs. You can design that function to perform whatever actions your scenario requires. The following code sample shows an event handler function that simply writes information about the event to the console.</span></span>

```js
function handleChange(event)
{
    return Excel.run(function(context){
        return context.sync()
            .then(function() {
                console.log("Change type of event: " + event.changeType);
                console.log("Address of event: " + event.address);
                console.log("Source of event: " + event.source);
            });
    }).catch(errorHandlerFunction);
}
```

## <a name="remove-an-event-handler"></a><span data-ttu-id="e7535-175">Remover um manipulador de eventos</span><span class="sxs-lookup"><span data-stu-id="e7535-175">Remove an event handler</span></span>

<span data-ttu-id="e7535-p110">O exemplo de código a seguir registra um manipulador de eventos para o evento `onSelectionChanged` na planilha **Sample** e define a função `handleSelectionChange` a executar quando o evento ocorrer. Também define a função `remove()` que pode ser chamada posteriormente para remover aquele manipulador de eventos.</span><span class="sxs-lookup"><span data-stu-id="e7535-p110">The following code sample registers an event handler for the `onSelectionChanged` event in the worksheet named **Sample** and defines the `handleSelectionChange` function that will run when the event occurs. It also defines the `remove()` function that can subsequently be called to remove that event handler.</span></span>

```js
var eventResult;

Excel.run(function (context) {
    var worksheet = context.workbook.worksheets.getItem("Sample");
    eventResult = worksheet.onSelectionChanged.add(handleSelectionChange);

    return context.sync()
        .then(function () {
            console.log("Event handler successfully registered for onSelectionChanged event in the worksheet.");
        });
}).catch(errorHandlerFunction);

function handleSelectionChange(event)
{
    return Excel.run(function(context){
        return context.sync()
            .then(function() {
                console.log("Address of current selection: " + event.address);
            });
    }).catch(errorHandlerFunction);
}

function remove() {
    return Excel.run(eventResult.context, function (context) {
        eventResult.remove();

        return context.sync()
            .then(function() {
                eventResult = null;
                console.log("Event handler successfully removed.");
            });
    }).catch(errorHandlerFunction);
}
```

## <a name="enable-and-disable-events"></a><span data-ttu-id="e7535-178">Habilitar e desabilitar eventos</span><span class="sxs-lookup"><span data-stu-id="e7535-178">Enable and disable events</span></span>

<span data-ttu-id="e7535-179">O desempenho de um suplemento pode ser melhorado desabilitando eventos.</span><span class="sxs-lookup"><span data-stu-id="e7535-179">The performance of an add-in may be improved by disabling events.</span></span>
<span data-ttu-id="e7535-180">Por exemplo, seu aplicativo pode não precisar receber eventos ou ele pode ignorar eventos durante a edições de lotes de várias entidades.</span><span class="sxs-lookup"><span data-stu-id="e7535-180">For example, your app might never need to receive events, or it could ignore events while performing batch-edits of multiple entities.</span></span>

<span data-ttu-id="e7535-181">Os eventos são habilitados ou desabilitados no nível [runtime](/javascript/api/excel/excel.runtime).</span><span class="sxs-lookup"><span data-stu-id="e7535-181">Events are enabled and disabled at the [runtime](/javascript/api/excel/excel.runtime) level.</span></span>
<span data-ttu-id="e7535-182">A `enableEvents` propriedade determina se os eventos são disparados e se seus manipuladores são ativados.</span><span class="sxs-lookup"><span data-stu-id="e7535-182">The `enableEvents` property determines if events are fired and their handlers are activated.</span></span>

<span data-ttu-id="e7535-183">O código a seguir mostra como ativar ou desativar os eventos.</span><span class="sxs-lookup"><span data-stu-id="e7535-183">The following code sample shows how to toggle events on and off.</span></span>

```js
Excel.run(function (context) {
    context.runtime.load("enableEvents");
    return context.sync()
        .then(function () {
            var eventBoolean = !context.runtime.enableEvents;
            context.runtime.enableEvents = eventBoolean;
            if (eventBoolean) {
                console.log("Events are currently on.");
            } else {
                console.log("Events are currently off.");
            }
        }).then(context.sync);
}).catch(errorHandlerFunction);
```

## <a name="see-also"></a><span data-ttu-id="e7535-184">Confira também</span><span class="sxs-lookup"><span data-stu-id="e7535-184">See also</span></span>

- [<span data-ttu-id="e7535-185">Conceitos fundamentais de programação com a API JavaScript do Excel</span><span class="sxs-lookup"><span data-stu-id="e7535-185">Fundamental programming concepts with the Excel JavaScript API</span></span>](excel-add-ins-core-concepts.md)
