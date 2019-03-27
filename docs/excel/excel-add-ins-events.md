---
title: Trabalhar com eventos usando a API JavaScript do Excel
description: ''
ms.date: 03/19/2019
localization_priority: Priority
ms.openlocfilehash: 08653a84c051709d16371d89672d3f7ebe2030b7
ms.sourcegitcommit: a2950492a2337de3180b713f5693fe82dbdd6a17
ms.translationtype: HT
ms.contentlocale: pt-BR
ms.lasthandoff: 03/27/2019
ms.locfileid: "30872015"
---
# <a name="work-with-events-using-the-excel-javascript-api"></a><span data-ttu-id="123b7-102">Trabalhar com eventos usando a API JavaScript do Excel</span><span class="sxs-lookup"><span data-stu-id="123b7-102">Work with Events using the Excel JavaScript API</span></span>

<span data-ttu-id="123b7-103">Este artigo descreve conceitos importantes relacionados ao trabalho com eventos no Excel e fornece exemplos de código que mostram como registrar manipuladores de eventos, lidar com eventos e remover manipuladores de eventos usando as APIs JavaScript do Excel.</span><span class="sxs-lookup"><span data-stu-id="123b7-103">This article describes important concepts related to working with events in Excel and provides code samples that show how to register event handlers, handle events, and remove event handlers using the Excel JavaScript API.</span></span> 

## <a name="events-in-excel"></a><span data-ttu-id="123b7-104">Eventos no Excel</span><span class="sxs-lookup"><span data-stu-id="123b7-104">Events in Excel</span></span>

<span data-ttu-id="123b7-p101">Sempre que ocorrerem certos tipos de alterações em uma pasta de trabalho do Excel, uma notificação do evento será ativada. Ao usar as APIs JavaScript do Excel, você pode registrar manipuladores de eventos que permitem que o suplemento execute automaticamente uma função designada quando ocorre um evento específico. Os eventos a seguir têm suporte no momento:</span><span class="sxs-lookup"><span data-stu-id="123b7-p101">Each time certain types of changes occur in an Excel workbook, an event notification fires. By using the Excel JavaScript API, you can register event handlers that allow your add-in to automatically run a designated function when a specific event occurs. The following events are currently supported.</span></span>

| <span data-ttu-id="123b7-108">Evento</span><span class="sxs-lookup"><span data-stu-id="123b7-108">Event</span></span> | <span data-ttu-id="123b7-109">Descrição</span><span class="sxs-lookup"><span data-stu-id="123b7-109">Description</span></span> | <span data-ttu-id="123b7-110">Objetos com suporte</span><span class="sxs-lookup"><span data-stu-id="123b7-110">Supported objects</span></span> |
|:---------------|:-------------|:-----------|
| `onAdded` | <span data-ttu-id="123b7-111">Evento que ocorre quando um objeto é adicionado.</span><span class="sxs-lookup"><span data-stu-id="123b7-111">Event that occurs when an object is added.</span></span> | <span data-ttu-id="123b7-112">[**ChartCollection**](/javascript/api/excel/excel.chartcollection), [**WorksheetCollection**](/javascript/api/excel/excel.worksheetcollection)</span><span class="sxs-lookup"><span data-stu-id="123b7-112">[**ChartCollection**](/javascript/api/excel/excel.chartcollection), [**WorksheetCollection**](/javascript/api/excel/excel.worksheetcollection)</span></span> |
| `onDeleted` | <span data-ttu-id="123b7-113">Evento que ocorre quando um objeto é excluído.</span><span class="sxs-lookup"><span data-stu-id="123b7-113">Event that occurs when an object is deleted.</span></span> | <span data-ttu-id="123b7-114">[**ChartCollection**](/javascript/api/excel/excel.chartcollection), [**WorksheetCollection**](/javascript/api/excel/excel.worksheetcollection)</span><span class="sxs-lookup"><span data-stu-id="123b7-114">[**ChartCollection**](/javascript/api/excel/excel.chartcollection), [**WorksheetCollection**](/javascript/api/excel/excel.worksheetcollection)</span></span> |
| `onActivated` | <span data-ttu-id="123b7-115">Evento que ocorre quando um objeto é ativado.</span><span class="sxs-lookup"><span data-stu-id="123b7-115">Event that occurs when an object is activated.</span></span> | <span data-ttu-id="123b7-116">[**Gráfico**](/javascript/api/excel/excel.chart), [**ChartCollection**](/javascript/api/excel/excel.chartcollection), [**WorksheetCollection**](/javascript/api/excel/excel.worksheetcollection), [**Planilha**](/javascript/api/excel/excel.worksheet)</span><span class="sxs-lookup"><span data-stu-id="123b7-116">[**Chart**](/javascript/api/excel/excel.chart), [**ChartCollection**](/javascript/api/excel/excel.chartcollection), [**WorksheetCollection**](/javascript/api/excel/excel.worksheetcollection), [**Worksheet**](/javascript/api/excel/excel.worksheet)</span></span> |
| `onDeactivated` | <span data-ttu-id="123b7-117">Evento que ocorre quando um objeto é desativado.</span><span class="sxs-lookup"><span data-stu-id="123b7-117">Event that occurs when an object is deactivated.</span></span> | <span data-ttu-id="123b7-118">[**Gráfico**](/javascript/api/excel/excel.chart), [**ChartCollection**](/javascript/api/excel/excel.chartcollection), [**WorksheetCollection**](/javascript/api/excel/excel.worksheetcollection), [**Planilha**](/javascript/api/excel/excel.worksheet)</span><span class="sxs-lookup"><span data-stu-id="123b7-118">[**Chart**](/javascript/api/excel/excel.chart), [**ChartCollection**](/javascript/api/excel/excel.chartcollection), [**WorksheetCollection**](/javascript/api/excel/excel.worksheetcollection), [**Worksheet**](/javascript/api/excel/excel.worksheet)</span></span> |
| `onCalculated` | <span data-ttu-id="123b7-119">Evento ocorre quando uma planilha terminou cálculo (ou todas as planilhas do conjunto terminaram).</span><span class="sxs-lookup"><span data-stu-id="123b7-119">Event that occurs when a worksheet has finished calculation (or all the worksheets of the collection have finished).</span></span> | <span data-ttu-id="123b7-120">[**WorksheetCollection**](/javascript/api/excel/excel.worksheetcollection), [**Planilha**](/javascript/api/excel/excel.worksheet)</span><span class="sxs-lookup"><span data-stu-id="123b7-120">[**WorksheetCollection**](/javascript/api/excel/excel.worksheetcollection), [**Worksheet**](/javascript/api/excel/excel.worksheet)</span></span> |
| `onChanged` | <span data-ttu-id="123b7-121">Evento que ocorre quando os dados das células são alterados.</span><span class="sxs-lookup"><span data-stu-id="123b7-121">Event that occurs when data within cells is changed.</span></span> | <span data-ttu-id="123b7-122">[**Planilha**](/javascript/api/excel/excel.worksheet), [**Tabela**](/javascript/api/excel/excel.table), [**TableCollection**](/javascript/api/excel/excel.tablecollection)</span><span class="sxs-lookup"><span data-stu-id="123b7-122">[**Worksheet**](/javascript/api/excel/excel.worksheet), [**Table**](/javascript/api/excel/excel.table), [**TableCollection**](/javascript/api/excel/excel.tablecollection)</span></span> |
| `onDataChanged` | <span data-ttu-id="123b7-123">Ocorre quando os dados ou a formatação dentro da associação são alterados.</span><span class="sxs-lookup"><span data-stu-id="123b7-123">Event that occurs when data or formatting within the binding is changed.</span></span> | [<span data-ttu-id="123b7-124">**Associação**</span><span class="sxs-lookup"><span data-stu-id="123b7-124">**Binding**</span></span>](/javascript/api/excel/excel.binding) |
| `onSelectionChanged` | <span data-ttu-id="123b7-125">Evento que ocorre quando uma célula ativa ou um intervalo selecionado são alterados.</span><span class="sxs-lookup"><span data-stu-id="123b7-125">Event that occurs when the active cell or selected range is changed.</span></span> | <span data-ttu-id="123b7-126">[**Planilha**](/javascript/api/excel/excel.worksheet), [ **Tabela**](/javascript/api/excel/excel.table), [ **Associação**](/javascript/api/excel/excel.binding)</span><span class="sxs-lookup"><span data-stu-id="123b7-126">[**Worksheet**](/javascript/api/excel/excel.worksheet), [**Table**](/javascript/api/excel/excel.table), [**Binding**](/javascript/api/excel/excel.binding)</span></span> |
| `onSettingsChanged` | <span data-ttu-id="123b7-127">O evento ocorre quando as Configurações no documento são alteradas.</span><span class="sxs-lookup"><span data-stu-id="123b7-127">Event that occurs when the Settings in the document are changed.</span></span> | [<span data-ttu-id="123b7-128">**SettingCollection**</span><span class="sxs-lookup"><span data-stu-id="123b7-128">**SettingCollection**</span></span>](/javascript/api/excel/excel.settingcollection) |

### <a name="event-triggers"></a><span data-ttu-id="123b7-129">Gatilhos de eventos</span><span class="sxs-lookup"><span data-stu-id="123b7-129">Event triggers</span></span>

<span data-ttu-id="123b7-130">Os eventos em uma pasta de trabalho do Excel podem ser acionados por:</span><span class="sxs-lookup"><span data-stu-id="123b7-130">Events within an Excel workbook can be triggered by:</span></span>

- <span data-ttu-id="123b7-131">Interação do usuário por meio da interface (IU) do Excel que altera a pasta de trabalho</span><span class="sxs-lookup"><span data-stu-id="123b7-131">User interaction via the Excel user interface (UI) that changes the workbook</span></span>
- <span data-ttu-id="123b7-132">Código de suplemento do Office (JavaScript) que altera a pasta de trabalho</span><span class="sxs-lookup"><span data-stu-id="123b7-132">Office Add-in (JavaScript) code that changes the workbook</span></span>
- <span data-ttu-id="123b7-133">Código de suplemento VBA (macro) que altera a pasta de trabalho</span><span class="sxs-lookup"><span data-stu-id="123b7-133">VBA add-in (macro) code that changes the workbook</span></span>

<span data-ttu-id="123b7-134">Todas as alterações que sejam compatíveis com o comportamento padrão do Excel acionarão eventos correspondentes em uma pasta de trabalho.</span><span class="sxs-lookup"><span data-stu-id="123b7-134">Any change that complies with default behavior of Excel will trigger the corresponding event(s) in a workbook.</span></span>

### <a name="lifecycle-of-an-event-handler"></a><span data-ttu-id="123b7-135">Ciclo de vida de um manipulador de eventos</span><span class="sxs-lookup"><span data-stu-id="123b7-135">Lifecycle of an event handler</span></span>

<span data-ttu-id="123b7-136">Um manipulador de eventos é criado quando um suplemento registra o manipulador de eventos.</span><span class="sxs-lookup"><span data-stu-id="123b7-136">An event handler is created when an add-in registers the event handler.</span></span> <span data-ttu-id="123b7-137">Ele é destruído quando o suplemento cancela o registro de manipulador de evento ou quando o suplemento é atualizado, recarregado ou fechado.</span><span class="sxs-lookup"><span data-stu-id="123b7-137">It is destroyed when the add-in unregisters the event handler or when the add-in is refreshed, reloaded, or closed.</span></span> <span data-ttu-id="123b7-138">Manipuladores de eventos não são mantidos como parte do arquivo do Excel ou em sessões do Excel Online.</span><span class="sxs-lookup"><span data-stu-id="123b7-138">Event handlers do not persist as part of the Excel file, or across sessions with Excel Online.</span></span>

> [!CAUTION]
> <span data-ttu-id="123b7-139">Quando um objeto ao qual os eventos são registrados é excluído (por exemplo, uma tabela com um `onChanged` evento registrado), o manipulador de eventos não disparará mais, mas permanecerá na memória até que o suplemento ou sessão do Excel atualize ou feche.</span><span class="sxs-lookup"><span data-stu-id="123b7-139">When an object to which events are registered is deleted (e.g., a table with an `onChanged` event registered), the event handler no longer triggers but remains in memory until the add-in or Excel session refreshes or closes.</span></span>

### <a name="events-and-coauthoring"></a><span data-ttu-id="123b7-140">Eventos e coautoria</span><span class="sxs-lookup"><span data-stu-id="123b7-140">Events and coauthoring</span></span>

<span data-ttu-id="123b7-p103">Com a [coautoria](co-authoring-in-excel-add-ins.md), várias pessoas podem trabalhar em conjunto e editar a mesma pasta de trabalho do Excel simultaneamente. Em eventos que podem ser disparados por um coautor, como `onChanged`, o objeto de **evento** respectivo conterá a propriedade **fonte** que indica se o evento foi acionado localmente pelo usuário atual (`event.source = Local`) ou pelo coautor remoto (`event.source = Remote`).</span><span class="sxs-lookup"><span data-stu-id="123b7-p103">With [coauthoring](co-authoring-in-excel-add-ins.md), multiple people can work together and edit the same Excel workbook simultaneously. For events that can be triggered by a coauthor, such as `onChanged`, the corresponding **Event** object will contain a **source** property that indicates whether the event was triggered locally by the current user (`event.source = Local`) or was triggered by the remote coauthor (`event.source = Remote`).</span></span>

## <a name="register-an-event-handler"></a><span data-ttu-id="123b7-143">Registrar um manipulador de eventos.</span><span class="sxs-lookup"><span data-stu-id="123b7-143">Register an event handler</span></span>

<span data-ttu-id="123b7-p104">O exemplo de código a seguir registra um manipulador de eventos para o evento `onChanged` na planilha **Sample**. O código especifica que, quando os dados forem alterados na planilha, a função `handleDataChange` deve ser executada.</span><span class="sxs-lookup"><span data-stu-id="123b7-p104">The following code sample registers an event handler for the `onChanged` event in the worksheet named **Sample**. The code specifies that when data changes in that worksheet, the `handleDataChange` function should run.</span></span>

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

## <a name="handle-an-event"></a><span data-ttu-id="123b7-146">Manipular um evento</span><span class="sxs-lookup"><span data-stu-id="123b7-146">Handle an event</span></span>

<span data-ttu-id="123b7-p105">Como mostrado no exemplo anterior, quando você registrar um manipulador de eventos, indica a função a ser executada quando o evento especificado ocorre. Você pode criar essa função para executar as ações que seu cenário exige. O exemplo de código a seguir mostra uma função de manipulador de eventos que simplesmente grava informações sobre o evento no console.</span><span class="sxs-lookup"><span data-stu-id="123b7-p105">As shown in the previous example, when you register an event handler, you indicate the function that should run when the specified event occurs. You can design that function to perform whatever actions your scenario requires. The following code sample shows an event handler function that simply writes information about the event to the console.</span></span> 

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

## <a name="remove-an-event-handler"></a><span data-ttu-id="123b7-150">Remover um manipulador de eventos</span><span class="sxs-lookup"><span data-stu-id="123b7-150">Remove an event handler</span></span>

<span data-ttu-id="123b7-p106">O exemplo de código a seguir registra um manipulador de eventos para o evento `onSelectionChanged` na planilha **Sample** e define a função `handleSelectionChange` a executar quando o evento ocorrer. Também define a função `remove()` que pode ser chamada posteriormente para remover aquele manipulador de eventos.</span><span class="sxs-lookup"><span data-stu-id="123b7-p106">The following code sample registers an event handler for the `onSelectionChanged` event in the worksheet named **Sample** and defines the `handleSelectionChange` function that will run when the event occurs. It also defines the `remove()` function that can subsequently be called to remove that event handler.</span></span>

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

## <a name="enable-and-disable-events"></a><span data-ttu-id="123b7-153">Habilitar e desabilitar eventos</span><span class="sxs-lookup"><span data-stu-id="123b7-153">Enable and disable events</span></span>

<span data-ttu-id="123b7-154">O desempenho de um suplemento pode ser melhorado desabilitando eventos.</span><span class="sxs-lookup"><span data-stu-id="123b7-154">The performance of an add-in may be improved by disabling events.</span></span> <span data-ttu-id="123b7-155">Por exemplo, seu aplicativo pode não precisar receber eventos ou ele pode ignorar eventos durante a edições de lotes de várias entidades.</span><span class="sxs-lookup"><span data-stu-id="123b7-155">For example, your app might never need to receive events, or it could ignore events while performing batch-edits of multiple entities.</span></span>

<span data-ttu-id="123b7-156">Os eventos são habilitados ou desabilitados no nível [runtime](/javascript/api/excel/excel.runtime).</span><span class="sxs-lookup"><span data-stu-id="123b7-156">Events are enabled and disabled at the [runtime](/javascript/api/excel/excel.runtime) level.</span></span>
<span data-ttu-id="123b7-157">A `enableEvents` propriedade determina se os eventos são disparados e se seus manipuladores são ativados.</span><span class="sxs-lookup"><span data-stu-id="123b7-157">The `enableEvents` property determines if events are fired and their handlers are activated.</span></span>

<span data-ttu-id="123b7-158">O código a seguir mostra como ativar ou desativar os eventos.</span><span class="sxs-lookup"><span data-stu-id="123b7-158">The following code sample shows how to toggle events on and off.</span></span>

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

## <a name="see-also"></a><span data-ttu-id="123b7-159">Confira também</span><span class="sxs-lookup"><span data-stu-id="123b7-159">See also</span></span>

- [<span data-ttu-id="123b7-160">Conceitos fundamentais de programação com a API JavaScript do Excel</span><span class="sxs-lookup"><span data-stu-id="123b7-160">Fundamental programming concepts with the Excel JavaScript API</span></span>](excel-add-ins-core-concepts.md)
