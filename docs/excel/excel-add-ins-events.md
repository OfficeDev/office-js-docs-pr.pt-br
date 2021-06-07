---
title: Trabalhar com eventos usando a API JavaScript do Excel
description: Uma lista de eventos para Excel JavaScript. Isso inclui informações sobre como usar manipuladores de eventos e os padrões associados.
ms.date: 06/02/2021
localization_priority: Normal
ms.openlocfilehash: 50d9c4c2a15a955f0a96c70464fa0165625ea6f8
ms.sourcegitcommit: 17b5a076375bc5dc3f91d3602daeb7535d67745d
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 06/06/2021
ms.locfileid: "52783502"
---
# <a name="work-with-events-using-the-excel-javascript-api"></a><span data-ttu-id="2887f-104">Trabalhar com eventos usando a API JavaScript do Excel</span><span class="sxs-lookup"><span data-stu-id="2887f-104">Work with Events using the Excel JavaScript API</span></span>

<span data-ttu-id="2887f-105">Este artigo descreve conceitos importantes relacionados ao trabalho com eventos no Excel e fornece exemplos de código que mostram como registrar manipuladores de eventos, lidar com eventos e remover manipuladores de eventos usando as APIs JavaScript do Excel.</span><span class="sxs-lookup"><span data-stu-id="2887f-105">This article describes important concepts related to working with events in Excel and provides code samples that show how to register event handlers, handle events, and remove event handlers using the Excel JavaScript API.</span></span>

## <a name="events-in-excel"></a><span data-ttu-id="2887f-106">Eventos no Excel</span><span class="sxs-lookup"><span data-stu-id="2887f-106">Events in Excel</span></span>

<span data-ttu-id="2887f-p102">Sempre que ocorrerem certos tipos de alterações em uma pasta de trabalho do Excel, uma notificação do evento será ativada. Ao usar as APIs JavaScript do Excel, você pode registrar manipuladores de eventos que permitem que o suplemento execute automaticamente uma função designada quando ocorre um evento específico. Os eventos a seguir têm suporte no momento:</span><span class="sxs-lookup"><span data-stu-id="2887f-p102">Each time certain types of changes occur in an Excel workbook, an event notification fires. By using the Excel JavaScript API, you can register event handlers that allow your add-in to automatically run a designated function when a specific event occurs. The following events are currently supported.</span></span>

| <span data-ttu-id="2887f-110">Evento</span><span class="sxs-lookup"><span data-stu-id="2887f-110">Event</span></span> | <span data-ttu-id="2887f-111">Descrição</span><span class="sxs-lookup"><span data-stu-id="2887f-111">Description</span></span> | <span data-ttu-id="2887f-112">Objetos com suporte</span><span class="sxs-lookup"><span data-stu-id="2887f-112">Supported objects</span></span> |
|:---------------|:-------------|:-----------|
| `onActivated` | <span data-ttu-id="2887f-113">Ocorre quando um objeto está ativado.</span><span class="sxs-lookup"><span data-stu-id="2887f-113">Occurs when an object is activated.</span></span> | <span data-ttu-id="2887f-114">[**Gráfico**](/javascript/api/excel/excel.chart#onactivated), [**ChartCollection**](/javascript/api/excel/excel.chartcollection#onactivated), [**Shape**](/javascript/api/excel/excel.shape#onactivated), [**Planilha**](/javascript/api/excel/excel.worksheet#onactivated), [**WorksheetCollection**](/javascript/api/excel/excel.worksheetcollection#onactivated)</span><span class="sxs-lookup"><span data-stu-id="2887f-114">[**Chart**](/javascript/api/excel/excel.chart#onactivated), [**ChartCollection**](/javascript/api/excel/excel.chartcollection#onactivated), [**Shape**](/javascript/api/excel/excel.shape#onactivated), [**Worksheet**](/javascript/api/excel/excel.worksheet#onactivated), [**WorksheetCollection**](/javascript/api/excel/excel.worksheetcollection#onactivated)</span></span> |
| `onAdded` | <span data-ttu-id="2887f-115">Ocorre quando um objeto é adicionado à coleção.</span><span class="sxs-lookup"><span data-stu-id="2887f-115">Occurs when an object is added to the collection.</span></span> | <span data-ttu-id="2887f-116">[**ChartCollection,**](/javascript/api/excel/excel.chartcollection#onadded) [**CommentCollection,**](/javascript/api/excel/excel.commentcollection#onadded) [**TableCollection,**](/javascript/api/excel/excel.tablecollection#onadded) [**WorksheetCollection**](/javascript/api/excel/excel.worksheetcollection#onadded)</span><span class="sxs-lookup"><span data-stu-id="2887f-116">[**ChartCollection**](/javascript/api/excel/excel.chartcollection#onadded), [**CommentCollection**](/javascript/api/excel/excel.commentcollection#onadded), [**TableCollection**](/javascript/api/excel/excel.tablecollection#onadded), [**WorksheetCollection**](/javascript/api/excel/excel.worksheetcollection#onadded)</span></span> |
| `onAutoSaveSettingChanged` | <span data-ttu-id="2887f-117">Ocorre quando a `autoSave` configuração é alterada na pasta de trabalho.</span><span class="sxs-lookup"><span data-stu-id="2887f-117">Occurs when the `autoSave` setting is changed on the workbook.</span></span> | [<span data-ttu-id="2887f-118">**Workbook**</span><span class="sxs-lookup"><span data-stu-id="2887f-118">**Workbook**</span></span>](/javascript/api/excel/excel.workbook#onautosavesettingchanged) |
| `onCalculated` | <span data-ttu-id="2887f-119">Ocorre quando uma planilha terminou um cálculo (ou todas as planilhas do conjunto terminaram).</span><span class="sxs-lookup"><span data-stu-id="2887f-119">Occurs when a worksheet has finished calculation (or all the worksheets of the collection have finished).</span></span> | <span data-ttu-id="2887f-120">[**WorksheetCollection**](/javascript/api/excel/excel.worksheet#oncalculated), [**Planilha**](/javascript/api/excel/excel.worksheetcollection#oncalculated)</span><span class="sxs-lookup"><span data-stu-id="2887f-120">[**Worksheet**](/javascript/api/excel/excel.worksheet#oncalculated), [**WorksheetCollection**](/javascript/api/excel/excel.worksheetcollection#oncalculated)</span></span> |
| `onChanged` | <span data-ttu-id="2887f-121">Ocorre quando os dados de células individuais ou comentários foram alterados.</span><span class="sxs-lookup"><span data-stu-id="2887f-121">Occurs when the data of individual cells or comments has changed.</span></span> | <span data-ttu-id="2887f-122">[**CommentCollection**](/javascript/api/excel/excel.commentcollection#onchanged), [**Table**](/javascript/api/excel/excel.table#onchanged), [**TableCollection**](/javascript/api/excel/excel.tablecollection#onchanged), [**Worksheet**](/javascript/api/excel/excel.worksheet#onchanged), [**WorksheetCollection**](/javascript/api/excel/excel.worksheetcollection#onchanged)</span><span class="sxs-lookup"><span data-stu-id="2887f-122">[**CommentCollection**](/javascript/api/excel/excel.commentcollection#onchanged), [**Table**](/javascript/api/excel/excel.table#onchanged), [**TableCollection**](/javascript/api/excel/excel.tablecollection#onchanged), [**Worksheet**](/javascript/api/excel/excel.worksheet#onchanged), [**WorksheetCollection**](/javascript/api/excel/excel.worksheetcollection#onchanged)</span></span> |
| `onColumnSorted` | <span data-ttu-id="2887f-123">Ocorre quando uma ou mais colunas são classificadas.</span><span class="sxs-lookup"><span data-stu-id="2887f-123">Occurs when one or more columns have been sorted.</span></span> <span data-ttu-id="2887f-124">Isso acontece como resultado de uma operação de classificação da esquerda para a direita.</span><span class="sxs-lookup"><span data-stu-id="2887f-124">This happens as the result of a left-to-right sort operation.</span></span> | <span data-ttu-id="2887f-125">[**Planilha**](/javascript/api/excel/excel.worksheet#oncolumnsorted), [**WorksheetCollection**](/javascript/api/excel/excel.worksheetcollection#oncolumnsorted)</span><span class="sxs-lookup"><span data-stu-id="2887f-125">[**Worksheet**](/javascript/api/excel/excel.worksheet#oncolumnsorted), [**WorksheetCollection**](/javascript/api/excel/excel.worksheetcollection#oncolumnsorted)</span></span> |
| `onDataChanged` | <span data-ttu-id="2887f-126">Ocorre quando os dados ou a formatação dentro da associação são alterados.</span><span class="sxs-lookup"><span data-stu-id="2887f-126">Occurs when data or formatting within the binding is changed.</span></span> | [<span data-ttu-id="2887f-127">**Associação**</span><span class="sxs-lookup"><span data-stu-id="2887f-127">**Binding**</span></span>](/javascript/api/excel/excel.binding#ondatachanged) |
| `onDeactivated` | <span data-ttu-id="2887f-128">Ocorre quando um objeto é desativado.</span><span class="sxs-lookup"><span data-stu-id="2887f-128">Occurs when an object is deactivated.</span></span> | <span data-ttu-id="2887f-129">[**Gráfico**](/javascript/api/excel/excel.chart#ondeactivated), [**ChartCollection**](/javascript/api/excel/excel.chartcollection#ondeactivated), [**Shape**](/javascript/api/excel/excel.shape#ondeactivated), [**Planilha**](/javascript/api/excel/excel.worksheet#ondeactivated), [**WorksheetCollection**](/javascript/api/excel/excel.worksheetcollection#ondeactivated)</span><span class="sxs-lookup"><span data-stu-id="2887f-129">[**Chart**](/javascript/api/excel/excel.chart#ondeactivated), [**ChartCollection**](/javascript/api/excel/excel.chartcollection#ondeactivated), [**Shape**](/javascript/api/excel/excel.shape#ondeactivated), [**Worksheet**](/javascript/api/excel/excel.worksheet#ondeactivated), [**WorksheetCollection**](/javascript/api/excel/excel.worksheetcollection#ondeactivated)</span></span> |
| `onDeleted` | <span data-ttu-id="2887f-130">Ocorre quando um objeto é excluído da coleção.</span><span class="sxs-lookup"><span data-stu-id="2887f-130">Occurs when an object is deleted from the collection.</span></span> | <span data-ttu-id="2887f-131">[**ChartCollection,**](/javascript/api/excel/excel.chartcollection#ondeleted) [**CommentCollection,**](/javascript/api/excel/excel.commentcollection#ondeleted) [**TableCollection,**](/javascript/api/excel/excel.tablecollection#ondeleted) [**WorksheetCollection**](/javascript/api/excel/excel.worksheetcollection#ondeleted)</span><span class="sxs-lookup"><span data-stu-id="2887f-131">[**ChartCollection**](/javascript/api/excel/excel.chartcollection#ondeleted), [**CommentCollection**](/javascript/api/excel/excel.commentcollection#ondeleted), [**TableCollection**](/javascript/api/excel/excel.tablecollection#ondeleted), [**WorksheetCollection**](/javascript/api/excel/excel.worksheetcollection#ondeleted)</span></span> |
| `onFormatChanged` | <span data-ttu-id="2887f-132">Ocorre quando o formato é alterado em uma planilha.</span><span class="sxs-lookup"><span data-stu-id="2887f-132">Occurs when the format is changed on a worksheet.</span></span> | <span data-ttu-id="2887f-133">[**Planilha**](/javascript/api/excel/excel.worksheet#onformatchanged), [**WorksheetCollection**](/javascript/api/excel/excel.worksheetcollection#onformatchanged)</span><span class="sxs-lookup"><span data-stu-id="2887f-133">[**Worksheet**](/javascript/api/excel/excel.worksheet#onformatchanged), [**WorksheetCollection**](/javascript/api/excel/excel.worksheetcollection#onformatchanged)</span></span> |
| `onRowSorted` | <span data-ttu-id="2887f-134">Ocorre quando uma ou mais linhas são classificadas.</span><span class="sxs-lookup"><span data-stu-id="2887f-134">Occurs when one or more rows have been sorted.</span></span> <span data-ttu-id="2887f-135">Isso ocorre como resultado de uma operação de classificação de cima para baixo.</span><span class="sxs-lookup"><span data-stu-id="2887f-135">This happens as the result of a top-to-bottom sort operation.</span></span> | <span data-ttu-id="2887f-136">[**Planilha**](/javascript/api/excel/excel.worksheet#onrowsorted), [**WorksheetCollection**](/javascript/api/excel/excel.worksheetcollection#onrowsorted)</span><span class="sxs-lookup"><span data-stu-id="2887f-136">[**Worksheet**](/javascript/api/excel/excel.worksheet#onrowsorted), [**WorksheetCollection**](/javascript/api/excel/excel.worksheetcollection#onrowsorted)</span></span> |
| `onSelectionChanged` | <span data-ttu-id="2887f-137">Ocorre quando uma célula ativa ou um intervalo selecionado são alterados.</span><span class="sxs-lookup"><span data-stu-id="2887f-137">Occurs when the active cell or selected range is changed.</span></span> | <span data-ttu-id="2887f-138">[**Binding**](/javascript/api/excel/excel.binding#onselectionchanged), [**Table**](/javascript/api/excel/excel.table#onselectionchanged), [**Workbook**](/javascript/api/excel/excel.workbook#onselectionchanged), [**Worksheet**](/javascript/api/excel/excel.worksheet#onselectionchanged), [**WorksheetCollection**](/javascript/api/excel/excel.worksheetcollection#onselectionchanged)</span><span class="sxs-lookup"><span data-stu-id="2887f-138">[**Binding**](/javascript/api/excel/excel.binding#onselectionchanged), [**Table**](/javascript/api/excel/excel.table#onselectionchanged), [**Workbook**](/javascript/api/excel/excel.workbook#onselectionchanged), [**Worksheet**](/javascript/api/excel/excel.worksheet#onselectionchanged), [**WorksheetCollection**](/javascript/api/excel/excel.worksheetcollection#onselectionchanged)</span></span> |
| `onRowHiddenChanged` | <span data-ttu-id="2887f-139">Ocorre quando o estado de linha oculta é alterado em uma planilha específica.</span><span class="sxs-lookup"><span data-stu-id="2887f-139">Occurs when the row-hidden state changes on a specific worksheet.</span></span> | <span data-ttu-id="2887f-140">[**Planilha**](/javascript/api/excel/excel.worksheet#onrowhiddenchanged), [**WorksheetCollection**](/javascript/api/excel/excel.worksheetcollection#onrowhiddenchanged)</span><span class="sxs-lookup"><span data-stu-id="2887f-140">[**Worksheet**](/javascript/api/excel/excel.worksheet#onrowhiddenchanged), [**WorksheetCollection**](/javascript/api/excel/excel.worksheetcollection#onrowhiddenchanged)</span></span> |
| `onSettingsChanged` | <span data-ttu-id="2887f-141">Ocorre quando as Configurações no documento são alteradas.</span><span class="sxs-lookup"><span data-stu-id="2887f-141">Occurs when the Settings in the document are changed.</span></span> | [<span data-ttu-id="2887f-142">**SettingCollection**</span><span class="sxs-lookup"><span data-stu-id="2887f-142">**SettingCollection**</span></span>](/javascript/api/excel/excel.settingcollection#onsettingschanged) |
| `onSingleClicked` | <span data-ttu-id="2887f-143">Acontece quando a operação é clicada/pressionada com o botão esquerdo do mouse ocorre na planilha.</span><span class="sxs-lookup"><span data-stu-id="2887f-143">Occurs when left-clicked/tapped action occurs in the worksheet.</span></span> | <span data-ttu-id="2887f-144">[**Planilha**](/javascript/api/excel/excel.worksheet#onsingleclicked), [**WorksheetCollection**](/javascript/api/excel/excel.worksheetcollection#onsingleclicked)</span><span class="sxs-lookup"><span data-stu-id="2887f-144">[**Worksheet**](/javascript/api/excel/excel.worksheet#onsingleclicked), [**WorksheetCollection**](/javascript/api/excel/excel.worksheetcollection#onsingleclicked)</span></span> |

### <a name="events-in-preview"></a><span data-ttu-id="2887f-145">Eventos no modo de visualização</span><span class="sxs-lookup"><span data-stu-id="2887f-145">Events in preview</span></span>

> [!NOTE]
> <span data-ttu-id="2887f-146">Os seguintes eventos estão disponíveis atualmente apenas na visualização pública.</span><span class="sxs-lookup"><span data-stu-id="2887f-146">The following events are currently available only in public preview.</span></span> [!INCLUDE [Information about using preview APIs](../includes/using-excel-preview-apis.md)]

| <span data-ttu-id="2887f-147">Evento</span><span class="sxs-lookup"><span data-stu-id="2887f-147">Event</span></span> | <span data-ttu-id="2887f-148">Descrição</span><span class="sxs-lookup"><span data-stu-id="2887f-148">Description</span></span> | <span data-ttu-id="2887f-149">Objetos com suporte</span><span class="sxs-lookup"><span data-stu-id="2887f-149">Supported objects</span></span> |
|:---------------|:-------------|:-----------|
| `onFiltered` | <span data-ttu-id="2887f-150">Ocorre quando um filtro é aplicado a um objeto.</span><span class="sxs-lookup"><span data-stu-id="2887f-150">Occurs when a filter is applied to an object.</span></span> | <span data-ttu-id="2887f-151">[**Tabela**](/javascript/api/excel/excel.table#onfiltered), [**TableCollection**](/javascript/api/excel/excel.tablecollection#onfiltered), [**Planilha**](/javascript/api/excel/excel.worksheet#onfiltered), [**WorksheetCollection**](/javascript/api/excel/excel.worksheetcollection#onfiltered)</span><span class="sxs-lookup"><span data-stu-id="2887f-151">[**Table**](/javascript/api/excel/excel.table#onfiltered), [**TableCollection**](/javascript/api/excel/excel.tablecollection#onfiltered), [**Worksheet**](/javascript/api/excel/excel.worksheet#onfiltered), [**WorksheetCollection**](/javascript/api/excel/excel.worksheetcollection#onfiltered)</span></span> |
| `onFormulaChanged` | <span data-ttu-id="2887f-152">Ocorre quando uma fórmula é alterada.</span><span class="sxs-lookup"><span data-stu-id="2887f-152">Occurs when a formula is changed.</span></span> | <span data-ttu-id="2887f-153">[**Planilha**](/javascript/api/excel/excel.worksheet#onFormulaChanged), [**WorksheetCollection**](/javascript/api/excel/excel.worksheetcollection#onFormulaChanged)</span><span class="sxs-lookup"><span data-stu-id="2887f-153">[**Worksheet**](/javascript/api/excel/excel.worksheet#onFormulaChanged), [**WorksheetCollection**](/javascript/api/excel/excel.worksheetcollection#onFormulaChanged)</span></span> |

### <a name="event-triggers"></a><span data-ttu-id="2887f-154">Gatilhos de eventos</span><span class="sxs-lookup"><span data-stu-id="2887f-154">Event triggers</span></span>

<span data-ttu-id="2887f-155">Os eventos em uma pasta de trabalho do Excel podem ser acionados por:</span><span class="sxs-lookup"><span data-stu-id="2887f-155">Events within an Excel workbook can be triggered by:</span></span>

- <span data-ttu-id="2887f-156">Interação do usuário por meio da interface (IU) do Excel que altera a pasta de trabalho</span><span class="sxs-lookup"><span data-stu-id="2887f-156">User interaction via the Excel user interface (UI) that changes the workbook</span></span>
- <span data-ttu-id="2887f-157">Código de suplemento do Office (JavaScript) que altera a pasta de trabalho</span><span class="sxs-lookup"><span data-stu-id="2887f-157">Office Add-in (JavaScript) code that changes the workbook</span></span>
- <span data-ttu-id="2887f-158">Código de suplemento VBA (macro) que altera a pasta de trabalho</span><span class="sxs-lookup"><span data-stu-id="2887f-158">VBA add-in (macro) code that changes the workbook</span></span>

<span data-ttu-id="2887f-159">Todas as alterações que sejam compatíveis com o comportamento padrão do Excel acionarão eventos correspondentes em uma pasta de trabalho.</span><span class="sxs-lookup"><span data-stu-id="2887f-159">Any change that complies with default behavior of Excel will trigger the corresponding event(s) in a workbook.</span></span>

### <a name="lifecycle-of-an-event-handler"></a><span data-ttu-id="2887f-160">Ciclo de vida de um manipulador de eventos</span><span class="sxs-lookup"><span data-stu-id="2887f-160">Lifecycle of an event handler</span></span>

<span data-ttu-id="2887f-161">Um manipulador de eventos é criado quando um suplemento registra o manipulador de eventos.</span><span class="sxs-lookup"><span data-stu-id="2887f-161">An event handler is created when an add-in registers the event handler.</span></span> <span data-ttu-id="2887f-162">Ele é destruído quando o suplemento cancela o registro de manipulador de evento ou quando o suplemento é atualizado, recarregado ou fechado.</span><span class="sxs-lookup"><span data-stu-id="2887f-162">It is destroyed when the add-in unregisters the event handler or when the add-in is refreshed, reloaded, or closed.</span></span> <span data-ttu-id="2887f-163">Manipuladores de eventos não são mantidos como parte do arquivo do Excel ou em sessões do Excel Online.</span><span class="sxs-lookup"><span data-stu-id="2887f-163">Event handlers do not persist as part of the Excel file, or across sessions with Excel on the web.</span></span>

> [!CAUTION]
> <span data-ttu-id="2887f-164">Quando um objeto ao qual os eventos são registrados é excluído (por exemplo, uma tabela com um `onChanged` evento registrado), o manipulador de eventos não disparará mais, mas permanecerá na memória até que o suplemento ou sessão do Excel atualize ou feche.</span><span class="sxs-lookup"><span data-stu-id="2887f-164">When an object to which events are registered is deleted (e.g., a table with an `onChanged` event registered), the event handler no longer triggers but remains in memory until the add-in or Excel session refreshes or closes.</span></span>

### <a name="events-and-coauthoring"></a><span data-ttu-id="2887f-165">Eventos e coautoria</span><span class="sxs-lookup"><span data-stu-id="2887f-165">Events and coauthoring</span></span>

<span data-ttu-id="2887f-p107">Com a [coautoria](co-authoring-in-excel-add-ins.md), várias pessoas podem trabalhar em conjunto e editar a mesma pasta de trabalho do Excel simultaneamente. Em eventos que podem ser disparados por um coautor, como `onChanged`, o objeto de **evento** respectivo conterá a propriedade **fonte** que indica se o evento foi acionado localmente pelo usuário atual (`event.source = Local`) ou pelo coautor remoto (`event.source = Remote`).</span><span class="sxs-lookup"><span data-stu-id="2887f-p107">With [coauthoring](co-authoring-in-excel-add-ins.md), multiple people can work together and edit the same Excel workbook simultaneously. For events that can be triggered by a coauthor, such as `onChanged`, the corresponding **Event** object will contain a **source** property that indicates whether the event was triggered locally by the current user (`event.source = Local`) or was triggered by the remote coauthor (`event.source = Remote`).</span></span>

## <a name="register-an-event-handler"></a><span data-ttu-id="2887f-168">Registrar um manipulador de eventos.</span><span class="sxs-lookup"><span data-stu-id="2887f-168">Register an event handler</span></span>

<span data-ttu-id="2887f-p108">O exemplo de código a seguir registra um manipulador de eventos para o evento `onChanged` na planilha **Sample**. O código especifica que, quando os dados forem alterados na planilha, a função `handleChange` deve ser executada.</span><span class="sxs-lookup"><span data-stu-id="2887f-p108">The following code sample registers an event handler for the `onChanged` event in the worksheet named **Sample**. The code specifies that when data changes in that worksheet, the `handleChange` function should run.</span></span>

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

## <a name="handle-an-event"></a><span data-ttu-id="2887f-171">Manipular um evento</span><span class="sxs-lookup"><span data-stu-id="2887f-171">Handle an event</span></span>

<span data-ttu-id="2887f-p109">Como mostrado no exemplo anterior, quando você registrar um manipulador de eventos, indica a função a ser executada quando o evento especificado ocorre. Você pode criar essa função para executar as ações que seu cenário exige. O exemplo de código a seguir mostra uma função de manipulador de eventos que simplesmente grava informações sobre o evento no console.</span><span class="sxs-lookup"><span data-stu-id="2887f-p109">As shown in the previous example, when you register an event handler, you indicate the function that should run when the specified event occurs. You can design that function to perform whatever actions your scenario requires. The following code sample shows an event handler function that simply writes information about the event to the console.</span></span>

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

## <a name="remove-an-event-handler"></a><span data-ttu-id="2887f-175">Remover um manipulador de eventos</span><span class="sxs-lookup"><span data-stu-id="2887f-175">Remove an event handler</span></span>

<span data-ttu-id="2887f-176">O exemplo de código a seguir registra um manipulador de eventos para o evento `onSelectionChanged` na planilha **Sample** e define a função `handleSelectionChange` a executar quando o evento ocorrer.</span><span class="sxs-lookup"><span data-stu-id="2887f-176">The following code sample registers an event handler for the `onSelectionChanged` event in the worksheet named **Sample** and defines the `handleSelectionChange` function that will run when the event occurs.</span></span> <span data-ttu-id="2887f-177">Também define a função `remove()` que pode ser chamada posteriormente para remover aquele manipulador de eventos.</span><span class="sxs-lookup"><span data-stu-id="2887f-177">It also defines the `remove()` function that can subsequently be called to remove that event handler.</span></span> <span data-ttu-id="2887f-178">Observe que o `RequestContext` usado para criar o manipulador de eventos é necessário para removê-lo.</span><span class="sxs-lookup"><span data-stu-id="2887f-178">Note that the `RequestContext` used to create the event handler is needed to remove it.</span></span> 

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

## <a name="enable-and-disable-events"></a><span data-ttu-id="2887f-179">Habilitar e desabilitar eventos</span><span class="sxs-lookup"><span data-stu-id="2887f-179">Enable and disable events</span></span>

<span data-ttu-id="2887f-180">O desempenho de um suplemento pode ser melhorado desabilitando eventos.</span><span class="sxs-lookup"><span data-stu-id="2887f-180">The performance of an add-in may be improved by disabling events.</span></span>
<span data-ttu-id="2887f-181">Por exemplo, seu aplicativo pode não precisar receber eventos ou ele pode ignorar eventos durante a edições de lotes de várias entidades.</span><span class="sxs-lookup"><span data-stu-id="2887f-181">For example, your app might never need to receive events, or it could ignore events while performing batch-edits of multiple entities.</span></span>

<span data-ttu-id="2887f-182">Os eventos são habilitados ou desabilitados no nível [runtime](/javascript/api/excel/excel.runtime).</span><span class="sxs-lookup"><span data-stu-id="2887f-182">Events are enabled and disabled at the [runtime](/javascript/api/excel/excel.runtime) level.</span></span>
<span data-ttu-id="2887f-183">A `enableEvents` propriedade determina se os eventos são disparados e se seus manipuladores são ativados.</span><span class="sxs-lookup"><span data-stu-id="2887f-183">The `enableEvents` property determines if events are fired and their handlers are activated.</span></span>

<span data-ttu-id="2887f-184">O código a seguir mostra como ativar ou desativar os eventos.</span><span class="sxs-lookup"><span data-stu-id="2887f-184">The following code sample shows how to toggle events on and off.</span></span>

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

## <a name="see-also"></a><span data-ttu-id="2887f-185">Confira também</span><span class="sxs-lookup"><span data-stu-id="2887f-185">See also</span></span>

- [<span data-ttu-id="2887f-186">Modelo de objeto JavaScript do Excel em Suplementos do Office</span><span class="sxs-lookup"><span data-stu-id="2887f-186">Excel JavaScript object model in Office Add-ins</span></span>](excel-add-ins-core-concepts.md)
