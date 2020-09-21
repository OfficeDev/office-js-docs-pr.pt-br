---
title: Trabalhar com eventos usando a API JavaScript do Excel
description: Uma lista de eventos para objetos JavaScript do Excel. Isso inclui informações sobre como usar manipuladores de eventos e os padrões associados.
ms.date: 09/15/2020
localization_priority: Normal
ms.openlocfilehash: 12e4f0a7ebb9c50bf6874106ef711ca3207cae46
ms.sourcegitcommit: 2479812e677d1a7337765fe8f1c8345061d4091a
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 09/19/2020
ms.locfileid: "48135211"
---
# <a name="work-with-events-using-the-excel-javascript-api"></a><span data-ttu-id="637ee-104">Trabalhar com eventos usando a API JavaScript do Excel</span><span class="sxs-lookup"><span data-stu-id="637ee-104">Work with Events using the Excel JavaScript API</span></span>

<span data-ttu-id="637ee-105">Este artigo descreve conceitos importantes relacionados ao trabalho com eventos no Excel e fornece exemplos de código que mostram como registrar manipuladores de eventos, lidar com eventos e remover manipuladores de eventos usando as APIs JavaScript do Excel.</span><span class="sxs-lookup"><span data-stu-id="637ee-105">This article describes important concepts related to working with events in Excel and provides code samples that show how to register event handlers, handle events, and remove event handlers using the Excel JavaScript API.</span></span>

## <a name="events-in-excel"></a><span data-ttu-id="637ee-106">Eventos no Excel</span><span class="sxs-lookup"><span data-stu-id="637ee-106">Events in Excel</span></span>

<span data-ttu-id="637ee-p102">Sempre que ocorrerem certos tipos de alterações em uma pasta de trabalho do Excel, uma notificação do evento será ativada. Ao usar as APIs JavaScript do Excel, você pode registrar manipuladores de eventos que permitem que o suplemento execute automaticamente uma função designada quando ocorre um evento específico. Os eventos a seguir têm suporte no momento:</span><span class="sxs-lookup"><span data-stu-id="637ee-p102">Each time certain types of changes occur in an Excel workbook, an event notification fires. By using the Excel JavaScript API, you can register event handlers that allow your add-in to automatically run a designated function when a specific event occurs. The following events are currently supported.</span></span>

| <span data-ttu-id="637ee-110">Evento</span><span class="sxs-lookup"><span data-stu-id="637ee-110">Event</span></span> | <span data-ttu-id="637ee-111">Descrição</span><span class="sxs-lookup"><span data-stu-id="637ee-111">Description</span></span> | <span data-ttu-id="637ee-112">Objetos com suporte</span><span class="sxs-lookup"><span data-stu-id="637ee-112">Supported objects</span></span> |
|:---------------|:-------------|:-----------|
| `onActivated` | <span data-ttu-id="637ee-113">Ocorre quando um objeto está ativado.</span><span class="sxs-lookup"><span data-stu-id="637ee-113">Occurs when an object is activated.</span></span> | <span data-ttu-id="637ee-114">[**Gráfico**](/javascript/api/excel/excel.chart#onactivated), [**ChartCollection**](/javascript/api/excel/excel.chartcollection#onactivated), [**Shape**](/javascript/api/excel/excel.shape#onactivated), [**Planilha**](/javascript/api/excel/excel.worksheet#onactivated), [**WorksheetCollection**](/javascript/api/excel/excel.worksheetcollection#onactivated)</span><span class="sxs-lookup"><span data-stu-id="637ee-114">[**Chart**](/javascript/api/excel/excel.chart#onactivated), [**ChartCollection**](/javascript/api/excel/excel.chartcollection#onactivated), [**Shape**](/javascript/api/excel/excel.shape#onactivated), [**Worksheet**](/javascript/api/excel/excel.worksheet#onactivated), [**WorksheetCollection**](/javascript/api/excel/excel.worksheetcollection#onactivated)</span></span> |
| `onAdded` | <span data-ttu-id="637ee-115">Ocorre quando um objeto é adicionado à coleção.</span><span class="sxs-lookup"><span data-stu-id="637ee-115">Occurs when an object is added to the collection.</span></span> | <span data-ttu-id="637ee-116">[**Chartcollection**](/javascript/api/excel/excel.chartcollection#onadded), [**commentcollection**](/javascript/api/excel/excel.commentcollection#onadded), [**TableCollection**](/javascript/api/excel/excel.tablecollection#onadded), [**worksheetcollection**](/javascript/api/excel/excel.worksheetcollection#onadded)</span><span class="sxs-lookup"><span data-stu-id="637ee-116">[**ChartCollection**](/javascript/api/excel/excel.chartcollection#onadded), [**CommentCollection**](/javascript/api/excel/excel.commentcollection#onadded), [**TableCollection**](/javascript/api/excel/excel.tablecollection#onadded), [**WorksheetCollection**](/javascript/api/excel/excel.worksheetcollection#onadded)</span></span> |
| `onAutoSaveSettingChanged` | <span data-ttu-id="637ee-117">Ocorre quando a `autoSave` configuração é alterada na pasta de trabalho.</span><span class="sxs-lookup"><span data-stu-id="637ee-117">Occurs when the `autoSave` setting is changed on the workbook.</span></span> | [<span data-ttu-id="637ee-118">**Workbook**</span><span class="sxs-lookup"><span data-stu-id="637ee-118">**Workbook**</span></span>](/javascript/api/excel/excel.workbook#onautosavesettingchanged) |
| `onCalculated` | <span data-ttu-id="637ee-119">Ocorre quando uma planilha terminou um cálculo (ou todas as planilhas do conjunto terminaram).</span><span class="sxs-lookup"><span data-stu-id="637ee-119">Occurs when a worksheet has finished calculation (or all the worksheets of the collection have finished).</span></span> | <span data-ttu-id="637ee-120">[**WorksheetCollection**](/javascript/api/excel/excel.worksheet#oncalculated), [**Planilha**](/javascript/api/excel/excel.worksheetcollection#oncalculated)</span><span class="sxs-lookup"><span data-stu-id="637ee-120">[**Worksheet**](/javascript/api/excel/excel.worksheet#oncalculated), [**WorksheetCollection**](/javascript/api/excel/excel.worksheetcollection#oncalculated)</span></span> |
| `onChanged` | <span data-ttu-id="637ee-121">Ocorre quando os dados de células ou Comentários individuais foram alterados.</span><span class="sxs-lookup"><span data-stu-id="637ee-121">Occurs when the data of individual cells or comments has changed.</span></span> | <span data-ttu-id="637ee-122">[**Commentcollection**](/javascript/api/excel/excel.commentcollection#onchanged), [**Table**](/javascript/api/excel/excel.table#onchanged), [**TableCollection**](/javascript/api/excel/excel.tablecollection#onchanged), [**Worksheet**](/javascript/api/excel/excel.worksheet#onchanged), [**worksheetcollection**](/javascript/api/excel/excel.worksheetcollection#onchanged)</span><span class="sxs-lookup"><span data-stu-id="637ee-122">[**CommentCollection**](/javascript/api/excel/excel.commentcollection#onchanged), [**Table**](/javascript/api/excel/excel.table#onchanged), [**TableCollection**](/javascript/api/excel/excel.tablecollection#onchanged), [**Worksheet**](/javascript/api/excel/excel.worksheet#onchanged), [**WorksheetCollection**](/javascript/api/excel/excel.worksheetcollection#onchanged)</span></span> |
| `onColumnSorted` | <span data-ttu-id="637ee-123">Ocorre quando uma ou mais colunas são classificadas.</span><span class="sxs-lookup"><span data-stu-id="637ee-123">Occurs when one or more columns have been sorted.</span></span> <span data-ttu-id="637ee-124">Isso acontece como resultado de uma operação de classificação da esquerda para a direita.</span><span class="sxs-lookup"><span data-stu-id="637ee-124">This happens as the result of a left-to-right sort operation.</span></span> | <span data-ttu-id="637ee-125">[**Planilha**](/javascript/api/excel/excel.worksheet#oncolumnsorted), [**WorksheetCollection**](/javascript/api/excel/excel.worksheetcollection#oncolumnsorted)</span><span class="sxs-lookup"><span data-stu-id="637ee-125">[**Worksheet**](/javascript/api/excel/excel.worksheet#oncolumnsorted), [**WorksheetCollection**](/javascript/api/excel/excel.worksheetcollection#oncolumnsorted)</span></span> |
| `onDataChanged` | <span data-ttu-id="637ee-126">Ocorre quando os dados ou a formatação dentro da associação são alterados.</span><span class="sxs-lookup"><span data-stu-id="637ee-126">Occurs when data or formatting within the binding is changed.</span></span> | [<span data-ttu-id="637ee-127">**Associação**</span><span class="sxs-lookup"><span data-stu-id="637ee-127">**Binding**</span></span>](/javascript/api/excel/excel.binding#ondatachanged) |
| `onDeactivated` | <span data-ttu-id="637ee-128">Ocorre quando um objeto é desativado.</span><span class="sxs-lookup"><span data-stu-id="637ee-128">Occurs when an object is deactivated.</span></span> | <span data-ttu-id="637ee-129">[**Gráfico**](/javascript/api/excel/excel.chart#ondeactivated), [**ChartCollection**](/javascript/api/excel/excel.chartcollection#ondeactivated), [**Shape**](/javascript/api/excel/excel.shape#ondeactivated), [**Planilha**](/javascript/api/excel/excel.worksheet#ondeactivated), [**WorksheetCollection**](/javascript/api/excel/excel.worksheetcollection#ondeactivated)</span><span class="sxs-lookup"><span data-stu-id="637ee-129">[**Chart**](/javascript/api/excel/excel.chart#ondeactivated), [**ChartCollection**](/javascript/api/excel/excel.chartcollection#ondeactivated), [**Shape**](/javascript/api/excel/excel.shape#ondeactivated), [**Worksheet**](/javascript/api/excel/excel.worksheet#ondeactivated), [**WorksheetCollection**](/javascript/api/excel/excel.worksheetcollection#ondeactivated)</span></span> |
| `onDeleted` | <span data-ttu-id="637ee-130">Ocorre quando um objeto é excluído da coleção.</span><span class="sxs-lookup"><span data-stu-id="637ee-130">Occurs when an object is deleted from the collection.</span></span> | <span data-ttu-id="637ee-131">[**Chartcollection**](/javascript/api/excel/excel.chartcollection#ondeleted), [**commentcollection**](/javascript/api/excel/excel.commentcollection#ondeleted), [**TableCollection**](/javascript/api/excel/excel.tablecollection#ondeleted), [**worksheetcollection**](/javascript/api/excel/excel.worksheetcollection#ondeleted)</span><span class="sxs-lookup"><span data-stu-id="637ee-131">[**ChartCollection**](/javascript/api/excel/excel.chartcollection#ondeleted), [**CommentCollection**](/javascript/api/excel/excel.commentcollection#ondeleted), [**TableCollection**](/javascript/api/excel/excel.tablecollection#ondeleted), [**WorksheetCollection**](/javascript/api/excel/excel.worksheetcollection#ondeleted)</span></span> |
| `onFormatChanged` | <span data-ttu-id="637ee-132">Ocorre quando o formato é alterado em uma planilha.</span><span class="sxs-lookup"><span data-stu-id="637ee-132">Occurs when the format is changed on a worksheet.</span></span> | <span data-ttu-id="637ee-133">[**Planilha**](/javascript/api/excel/excel.worksheet#onformatchanged), [**WorksheetCollection**](/javascript/api/excel/excel.worksheetcollection#onformatchanged)</span><span class="sxs-lookup"><span data-stu-id="637ee-133">[**Worksheet**](/javascript/api/excel/excel.worksheet#onformatchanged), [**WorksheetCollection**](/javascript/api/excel/excel.worksheetcollection#onformatchanged)</span></span> |
| `onRowSorted` | <span data-ttu-id="637ee-134">Ocorre quando uma ou mais linhas são classificadas.</span><span class="sxs-lookup"><span data-stu-id="637ee-134">Occurs when one or more rows have been sorted.</span></span> <span data-ttu-id="637ee-135">Isso ocorre como resultado de uma operação de classificação de cima para baixo.</span><span class="sxs-lookup"><span data-stu-id="637ee-135">This happens as the result of a top-to-bottom sort operation.</span></span> | <span data-ttu-id="637ee-136">[**Planilha**](/javascript/api/excel/excel.worksheet#onrowsorted), [**WorksheetCollection**](/javascript/api/excel/excel.worksheetcollection#onrowsorted)</span><span class="sxs-lookup"><span data-stu-id="637ee-136">[**Worksheet**](/javascript/api/excel/excel.worksheet#onrowsorted), [**WorksheetCollection**](/javascript/api/excel/excel.worksheetcollection#onrowsorted)</span></span> |
| `onSelectionChanged` | <span data-ttu-id="637ee-137">Ocorre quando uma célula ativa ou um intervalo selecionado são alterados.</span><span class="sxs-lookup"><span data-stu-id="637ee-137">Occurs when the active cell or selected range is changed.</span></span> | <span data-ttu-id="637ee-138">[**Associação**](/javascript/api/excel/excel.binding#onselectionchanged), [**tabela**](/javascript/api/excel/excel.table#onselectionchanged), [**pasta de trabalho**](/javascript/api/excel/excel.workbook#onselectionchanged), [**planilha**](/javascript/api/excel/excel.worksheet#onselectionchanged), [**planilhacollection**](/javascript/api/excel/excel.worksheetcollection#onselectionchanged)</span><span class="sxs-lookup"><span data-stu-id="637ee-138">[**Binding**](/javascript/api/excel/excel.binding#onselectionchanged), [**Table**](/javascript/api/excel/excel.table#onselectionchanged), [**Workbook**](/javascript/api/excel/excel.workbook#onselectionchanged), [**Worksheet**](/javascript/api/excel/excel.worksheet#onselectionchanged), [**WorksheetCollection**](/javascript/api/excel/excel.worksheetcollection#onselectionchanged)</span></span> |
| `onRowHiddenChanged` | <span data-ttu-id="637ee-139">Ocorre quando o estado de linha oculta é alterado em uma planilha específica.</span><span class="sxs-lookup"><span data-stu-id="637ee-139">Occurs when the row-hidden state changes on a specific worksheet.</span></span> | <span data-ttu-id="637ee-140">[**Planilha**](/javascript/api/excel/excel.worksheet#onrowhiddenchanged), [**WorksheetCollection**](/javascript/api/excel/excel.worksheetcollection#onrowhiddenchanged)</span><span class="sxs-lookup"><span data-stu-id="637ee-140">[**Worksheet**](/javascript/api/excel/excel.worksheet#onrowhiddenchanged), [**WorksheetCollection**](/javascript/api/excel/excel.worksheetcollection#onrowhiddenchanged)</span></span> |
| `onSettingsChanged` | <span data-ttu-id="637ee-141">Ocorre quando as Configurações no documento são alteradas.</span><span class="sxs-lookup"><span data-stu-id="637ee-141">Occurs when the Settings in the document are changed.</span></span> | [<span data-ttu-id="637ee-142">**SettingCollection**</span><span class="sxs-lookup"><span data-stu-id="637ee-142">**SettingCollection**</span></span>](/javascript/api/excel/excel.settingcollection#onsettingschanged) |
| `onSingleClicked` | <span data-ttu-id="637ee-143">Acontece quando a operação é clicada/pressionada com o botão esquerdo do mouse ocorre na planilha.</span><span class="sxs-lookup"><span data-stu-id="637ee-143">Occurs when left-clicked/tapped action occurs in the worksheet.</span></span> | <span data-ttu-id="637ee-144">[**Planilha**](/javascript/api/excel/excel.worksheet#onsingleclicked), [**WorksheetCollection**](/javascript/api/excel/excel.worksheetcollection#onsingleclicked)</span><span class="sxs-lookup"><span data-stu-id="637ee-144">[**Worksheet**](/javascript/api/excel/excel.worksheet#onsingleclicked), [**WorksheetCollection**](/javascript/api/excel/excel.worksheetcollection#onsingleclicked)</span></span> |

### <a name="events-in-preview"></a><span data-ttu-id="637ee-145">Eventos no modo de visualização</span><span class="sxs-lookup"><span data-stu-id="637ee-145">Events in preview</span></span>

> [!NOTE]
> <span data-ttu-id="637ee-146">Os seguintes eventos estão disponíveis atualmente apenas na visualização pública.</span><span class="sxs-lookup"><span data-stu-id="637ee-146">The following events are currently available only in public preview.</span></span> [!INCLUDE [Information about using preview APIs](../includes/using-excel-preview-apis.md)]

| <span data-ttu-id="637ee-147">Evento</span><span class="sxs-lookup"><span data-stu-id="637ee-147">Event</span></span> | <span data-ttu-id="637ee-148">Descrição</span><span class="sxs-lookup"><span data-stu-id="637ee-148">Description</span></span> | <span data-ttu-id="637ee-149">Objetos com suporte</span><span class="sxs-lookup"><span data-stu-id="637ee-149">Supported objects</span></span> |
|:---------------|:-------------|:-----------|
| `onFiltered` | <span data-ttu-id="637ee-150">Ocorre quando um filtro é aplicado a um objeto.</span><span class="sxs-lookup"><span data-stu-id="637ee-150">Occurs when a filter is applied to an object.</span></span> | <span data-ttu-id="637ee-151">[**Tabela**](/javascript/api/excel/excel.table#onfiltered), [**TableCollection**](/javascript/api/excel/excel.tablecollection#onfiltered), [**Planilha**](/javascript/api/excel/excel.worksheet#onfiltered), [**WorksheetCollection**](/javascript/api/excel/excel.worksheetcollection#onfiltered)</span><span class="sxs-lookup"><span data-stu-id="637ee-151">[**Table**](/javascript/api/excel/excel.table#onfiltered), [**TableCollection**](/javascript/api/excel/excel.tablecollection#onfiltered), [**Worksheet**](/javascript/api/excel/excel.worksheet#onfiltered), [**WorksheetCollection**](/javascript/api/excel/excel.worksheetcollection#onfiltered)</span></span> |

### <a name="event-triggers"></a><span data-ttu-id="637ee-152">Gatilhos de eventos</span><span class="sxs-lookup"><span data-stu-id="637ee-152">Event triggers</span></span>

<span data-ttu-id="637ee-153">Os eventos em uma pasta de trabalho do Excel podem ser acionados por:</span><span class="sxs-lookup"><span data-stu-id="637ee-153">Events within an Excel workbook can be triggered by:</span></span>

- <span data-ttu-id="637ee-154">Interação do usuário por meio da interface (IU) do Excel que altera a pasta de trabalho</span><span class="sxs-lookup"><span data-stu-id="637ee-154">User interaction via the Excel user interface (UI) that changes the workbook</span></span>
- <span data-ttu-id="637ee-155">Código de suplemento do Office (JavaScript) que altera a pasta de trabalho</span><span class="sxs-lookup"><span data-stu-id="637ee-155">Office Add-in (JavaScript) code that changes the workbook</span></span>
- <span data-ttu-id="637ee-156">Código de suplemento VBA (macro) que altera a pasta de trabalho</span><span class="sxs-lookup"><span data-stu-id="637ee-156">VBA add-in (macro) code that changes the workbook</span></span>

<span data-ttu-id="637ee-157">Todas as alterações que sejam compatíveis com o comportamento padrão do Excel acionarão eventos correspondentes em uma pasta de trabalho.</span><span class="sxs-lookup"><span data-stu-id="637ee-157">Any change that complies with default behavior of Excel will trigger the corresponding event(s) in a workbook.</span></span>

### <a name="lifecycle-of-an-event-handler"></a><span data-ttu-id="637ee-158">Ciclo de vida de um manipulador de eventos</span><span class="sxs-lookup"><span data-stu-id="637ee-158">Lifecycle of an event handler</span></span>

<span data-ttu-id="637ee-159">Um manipulador de eventos é criado quando um suplemento registra o manipulador de eventos.</span><span class="sxs-lookup"><span data-stu-id="637ee-159">An event handler is created when an add-in registers the event handler.</span></span> <span data-ttu-id="637ee-160">Ele é destruído quando o suplemento cancela o registro de manipulador de evento ou quando o suplemento é atualizado, recarregado ou fechado.</span><span class="sxs-lookup"><span data-stu-id="637ee-160">It is destroyed when the add-in unregisters the event handler or when the add-in is refreshed, reloaded, or closed.</span></span> <span data-ttu-id="637ee-161">Manipuladores de eventos não são mantidos como parte do arquivo do Excel ou em sessões do Excel Online.</span><span class="sxs-lookup"><span data-stu-id="637ee-161">Event handlers do not persist as part of the Excel file, or across sessions with Excel on the web.</span></span>

> [!CAUTION]
> <span data-ttu-id="637ee-162">Quando um objeto ao qual os eventos são registrados é excluído (por exemplo, uma tabela com um `onChanged` evento registrado), o manipulador de eventos não disparará mais, mas permanecerá na memória até que o suplemento ou sessão do Excel atualize ou feche.</span><span class="sxs-lookup"><span data-stu-id="637ee-162">When an object to which events are registered is deleted (e.g., a table with an `onChanged` event registered), the event handler no longer triggers but remains in memory until the add-in or Excel session refreshes or closes.</span></span>

### <a name="events-and-coauthoring"></a><span data-ttu-id="637ee-163">Eventos e coautoria</span><span class="sxs-lookup"><span data-stu-id="637ee-163">Events and coauthoring</span></span>

<span data-ttu-id="637ee-p107">Com a [coautoria](co-authoring-in-excel-add-ins.md), várias pessoas podem trabalhar em conjunto e editar a mesma pasta de trabalho do Excel simultaneamente. Em eventos que podem ser disparados por um coautor, como `onChanged`, o objeto de **evento** respectivo conterá a propriedade **fonte** que indica se o evento foi acionado localmente pelo usuário atual (`event.source = Local`) ou pelo coautor remoto (`event.source = Remote`).</span><span class="sxs-lookup"><span data-stu-id="637ee-p107">With [coauthoring](co-authoring-in-excel-add-ins.md), multiple people can work together and edit the same Excel workbook simultaneously. For events that can be triggered by a coauthor, such as `onChanged`, the corresponding **Event** object will contain a **source** property that indicates whether the event was triggered locally by the current user (`event.source = Local`) or was triggered by the remote coauthor (`event.source = Remote`).</span></span>

## <a name="register-an-event-handler"></a><span data-ttu-id="637ee-166">Registrar um manipulador de eventos.</span><span class="sxs-lookup"><span data-stu-id="637ee-166">Register an event handler</span></span>

<span data-ttu-id="637ee-p108">O exemplo de código a seguir registra um manipulador de eventos para o evento `onChanged` na planilha **Sample**. O código especifica que, quando os dados forem alterados na planilha, a função `handleDataChange` deve ser executada.</span><span class="sxs-lookup"><span data-stu-id="637ee-p108">The following code sample registers an event handler for the `onChanged` event in the worksheet named **Sample**. The code specifies that when data changes in that worksheet, the `handleDataChange` function should run.</span></span>

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

## <a name="handle-an-event"></a><span data-ttu-id="637ee-169">Manipular um evento</span><span class="sxs-lookup"><span data-stu-id="637ee-169">Handle an event</span></span>

<span data-ttu-id="637ee-p109">Como mostrado no exemplo anterior, quando você registrar um manipulador de eventos, indica a função a ser executada quando o evento especificado ocorre. Você pode criar essa função para executar as ações que seu cenário exige. O exemplo de código a seguir mostra uma função de manipulador de eventos que simplesmente grava informações sobre o evento no console.</span><span class="sxs-lookup"><span data-stu-id="637ee-p109">As shown in the previous example, when you register an event handler, you indicate the function that should run when the specified event occurs. You can design that function to perform whatever actions your scenario requires. The following code sample shows an event handler function that simply writes information about the event to the console.</span></span>

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

## <a name="remove-an-event-handler"></a><span data-ttu-id="637ee-173">Remover um manipulador de eventos</span><span class="sxs-lookup"><span data-stu-id="637ee-173">Remove an event handler</span></span>

<span data-ttu-id="637ee-174">O exemplo de código a seguir registra um manipulador de eventos para o evento `onSelectionChanged` na planilha **Sample** e define a função `handleSelectionChange` a executar quando o evento ocorrer.</span><span class="sxs-lookup"><span data-stu-id="637ee-174">The following code sample registers an event handler for the `onSelectionChanged` event in the worksheet named **Sample** and defines the `handleSelectionChange` function that will run when the event occurs.</span></span> <span data-ttu-id="637ee-175">Também define a função `remove()` que pode ser chamada posteriormente para remover aquele manipulador de eventos.</span><span class="sxs-lookup"><span data-stu-id="637ee-175">It also defines the `remove()` function that can subsequently be called to remove that event handler.</span></span> <span data-ttu-id="637ee-176">Observe que o `RequestContext` manipulador de eventos usado para criar o é necessário para removê-lo.</span><span class="sxs-lookup"><span data-stu-id="637ee-176">Note that the `RequestContext` used to create the event handler is needed to remove it.</span></span> 

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

## <a name="enable-and-disable-events"></a><span data-ttu-id="637ee-177">Habilitar e desabilitar eventos</span><span class="sxs-lookup"><span data-stu-id="637ee-177">Enable and disable events</span></span>

<span data-ttu-id="637ee-178">O desempenho de um suplemento pode ser melhorado desabilitando eventos.</span><span class="sxs-lookup"><span data-stu-id="637ee-178">The performance of an add-in may be improved by disabling events.</span></span>
<span data-ttu-id="637ee-179">Por exemplo, seu aplicativo pode não precisar receber eventos ou ele pode ignorar eventos durante a edições de lotes de várias entidades.</span><span class="sxs-lookup"><span data-stu-id="637ee-179">For example, your app might never need to receive events, or it could ignore events while performing batch-edits of multiple entities.</span></span>

<span data-ttu-id="637ee-180">Os eventos são habilitados ou desabilitados no nível [runtime](/javascript/api/excel/excel.runtime).</span><span class="sxs-lookup"><span data-stu-id="637ee-180">Events are enabled and disabled at the [runtime](/javascript/api/excel/excel.runtime) level.</span></span>
<span data-ttu-id="637ee-181">A `enableEvents` propriedade determina se os eventos são disparados e se seus manipuladores são ativados.</span><span class="sxs-lookup"><span data-stu-id="637ee-181">The `enableEvents` property determines if events are fired and their handlers are activated.</span></span>

<span data-ttu-id="637ee-182">O código a seguir mostra como ativar ou desativar os eventos.</span><span class="sxs-lookup"><span data-stu-id="637ee-182">The following code sample shows how to toggle events on and off.</span></span>

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

## <a name="see-also"></a><span data-ttu-id="637ee-183">Confira também</span><span class="sxs-lookup"><span data-stu-id="637ee-183">See also</span></span>

- [<span data-ttu-id="637ee-184">Modelo de objeto do JavaScript do Excel em suplementos do Office</span><span class="sxs-lookup"><span data-stu-id="637ee-184">Excel JavaScript object model in Office Add-ins</span></span>](excel-add-ins-core-concepts.md)
