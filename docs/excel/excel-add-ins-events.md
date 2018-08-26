---
title: Trabalhar com eventos usando a API JavaScript do Excel
description: ''
ms.date: 05/25/2018
ms.openlocfilehash: 3d94a36a60220b856795b8d0abf5387fcb8c1bad
ms.sourcegitcommit: e1c92ba882e6eb03a165867c6021a6aa742aa310
ms.translationtype: HT
ms.contentlocale: pt-BR
ms.lasthandoff: 08/20/2018
ms.locfileid: "22925623"
---
# <a name="work-with-events-using-the-excel-javascript-api"></a><span data-ttu-id="330cd-102">Trabalhar com eventos usando a API JavaScript do Excel</span><span class="sxs-lookup"><span data-stu-id="330cd-102">Work with Events using the Excel JavaScript API</span></span> 

<span data-ttu-id="330cd-103">Este artigo descreve conceitos importantes relacionados ao trabalho com eventos no Excel e fornece exemplos de código que mostram como registrar manipuladores de eventos, lidar com eventos e remover manipuladores de eventos usando as APIs JavaScript do Excel.</span><span class="sxs-lookup"><span data-stu-id="330cd-103">This article describes important concepts related to working with events in Excel and provides code samples that show how to register event handlers, handle events, and remove event handlers using the Excel JavaScript API.</span></span> 

## <a name="events-in-excel"></a><span data-ttu-id="330cd-104">Eventos no Excel</span><span class="sxs-lookup"><span data-stu-id="330cd-104">Events in Excel</span></span>

<span data-ttu-id="330cd-105">Sempre que ocorrerem certos tipos de alterações em uma pasta de trabalho do Excel, uma notificação do evento será ativada.</span><span class="sxs-lookup"><span data-stu-id="330cd-105">Each time certain types of changes occur in an Excel workbook, an event notification fires.</span></span> <span data-ttu-id="330cd-106">Ao usar as APIs JavaScript do Excel, você pode registrar manipuladores de eventos que permitem que o suplemento execute automaticamente uma função designada quando ocorre um evento específico.</span><span class="sxs-lookup"><span data-stu-id="330cd-106">By using the Excel JavaScript API, you can register event handlers that allow your add-in to automatically run a designated function when a specific event occurs.</span></span> <span data-ttu-id="330cd-107">Os eventos a seguir têm suporte no momento:</span><span class="sxs-lookup"><span data-stu-id="330cd-107">The following events are currently supported.</span></span>

| <span data-ttu-id="330cd-108">Evento</span><span class="sxs-lookup"><span data-stu-id="330cd-108">Event</span></span> | <span data-ttu-id="330cd-109">Descrição</span><span class="sxs-lookup"><span data-stu-id="330cd-109">Description</span></span> | <span data-ttu-id="330cd-110">Objetos com suporte</span><span class="sxs-lookup"><span data-stu-id="330cd-110">Supported objects</span></span> |
|:---------------|:-------------|:-----------|
| `onAdded` | <span data-ttu-id="330cd-111">Evento que ocorre quando um objeto é adicionado.</span><span class="sxs-lookup"><span data-stu-id="330cd-111">Event that occurs when an object is added.</span></span> | [<span data-ttu-id="330cd-112">**WorksheetCollection**</span><span class="sxs-lookup"><span data-stu-id="330cd-112">**WorksheetCollection**</span></span>](https://dev.office.com/reference/add-ins/excel/worksheetcollection) |
| `onDeleted` | <span data-ttu-id="330cd-113">Evento que ocorre quando um objeto é excluído.</span><span class="sxs-lookup"><span data-stu-id="330cd-113">Event that occurs when an object is deleted.</span></span> | [<span data-ttu-id="330cd-114">**WorksheetCollection**</span><span class="sxs-lookup"><span data-stu-id="330cd-114">**WorksheetCollection**</span></span>](https://dev.office.com/reference/add-ins/excel/worksheetcollection) |
| `onActivated` | <span data-ttu-id="330cd-115">Evento que ocorre quando um objeto é ativado.</span><span class="sxs-lookup"><span data-stu-id="330cd-115">Event that occurs when an object is activated.</span></span> | <span data-ttu-id="330cd-116">[**WorksheetCollection**](https://dev.office.com/reference/add-ins/excel/worksheetcollection), [**Planilha**](https://dev.office.com/reference/add-ins/excel/worksheet)</span><span class="sxs-lookup"><span data-stu-id="330cd-116">**WorksheetCollection**, [Worksheet](https://dev.office.com/reference/add-ins/excel/worksheetcollection)</span></span> |
| `onDeactivated` | <span data-ttu-id="330cd-117">Evento que ocorre quando um objeto é desativado.</span><span class="sxs-lookup"><span data-stu-id="330cd-117">Event that occurs when an object is deactivated.</span></span> | <span data-ttu-id="330cd-118">[**WorksheetCollection**](https://dev.office.com/reference/add-ins/excel/worksheetcollection), [**Planilha**](https://dev.office.com/reference/add-ins/excel/worksheet)</span><span class="sxs-lookup"><span data-stu-id="330cd-118">**WorksheetCollection**, [Worksheet](https://dev.office.com/reference/add-ins/excel/worksheetcollection)</span></span> |
| `onChanged` | <span data-ttu-id="330cd-119">Evento que ocorre quando os dados de células são alterados.</span><span class="sxs-lookup"><span data-stu-id="330cd-119">Event that occurs when data within cells is changed.</span></span> | <span data-ttu-id="330cd-120">[**Planilha**](https://dev.office.com/reference/add-ins/excel/worksheet), [**Tabela**](https://dev.office.com/reference/add-ins/excel/table), [**TableCollection**](https://dev.office.com/reference/add-ins/excel/tablecollection)</span><span class="sxs-lookup"><span data-stu-id="330cd-120">[**Worksheet**](https://dev.office.com/reference/add-ins/excel/worksheet), [**Table**](https://dev.office.com/reference/add-ins/excel/table), [**TableCollection**](https://dev.office.com/reference/add-ins/excel/tablecollection)</span></span> |
| `onDataChanged` | <span data-ttu-id="330cd-121">Evento que ocorre quando os dados ou a formatação na associação são alterados.</span><span class="sxs-lookup"><span data-stu-id="330cd-121">Occurs when data or formatting within the binding is changed.</span></span> | [<span data-ttu-id="330cd-122">**Associação**</span><span class="sxs-lookup"><span data-stu-id="330cd-122">**Binding**</span></span>](https://dev.office.com/reference/add-ins/excel/binding) |
| `onSelectionChanged` | <span data-ttu-id="330cd-123">Evento que ocorre quando uma célula ativa ou um intervalo selecionado são alterados.</span><span class="sxs-lookup"><span data-stu-id="330cd-123">Event that occurs when the active cell or selected range is changed.</span></span> | <span data-ttu-id="330cd-124">[**Planilha**](https://dev.office.com/reference/add-ins/excel/worksheet), [**Tabela**](https://dev.office.com/reference/add-ins/excel/table), [**Associação**](https://dev.office.com/reference/add-ins/excel/binding)</span><span class="sxs-lookup"><span data-stu-id="330cd-124">**Worksheet**, [Table](https://dev.office.com/reference/add-ins/excel/worksheet), **Binding**</span></span> |
| `onSettingsChanged` | <span data-ttu-id="330cd-125">Evento que ocorre quando as Configurações no documento são alteradas.</span><span class="sxs-lookup"><span data-stu-id="330cd-125">Occurs when the Settings in the document are changed.</span></span> | [<span data-ttu-id="330cd-126">**SettingCollection**</span><span class="sxs-lookup"><span data-stu-id="330cd-126">**SettingCollection**</span></span>](https://dev.office.com/reference/add-ins/excel/settingcollection) |

## <a name="preview-beta-events-in-excel"></a><span data-ttu-id="330cd-127">Visualizar eventos (beta) no Excel</span><span class="sxs-lookup"><span data-stu-id="330cd-127">Preview (Beta) Events in Excel</span></span>

> [!NOTE]
> <span data-ttu-id="330cd-128">Esses eventos estão atualmente disponíveis apenas na versão prévia pública (beta).</span><span class="sxs-lookup"><span data-stu-id="330cd-128">This sample uses APIs that are currently available only in public preview (beta).</span></span> <span data-ttu-id="330cd-129">Para usar esses recursos, você deve usar a biblioteca beta do CDN do Office.js: https://appsforoffice.microsoft.com/lib/beta/hosted/office.js.</span><span class="sxs-lookup"><span data-stu-id="330cd-129">To use these features, you must use the beta library of the Office.js CDN: https://appsforoffice.microsoft.com/lib/beta/hosted/office.js.</span></span>

| <span data-ttu-id="330cd-130">Evento</span><span class="sxs-lookup"><span data-stu-id="330cd-130">Event</span></span> | <span data-ttu-id="330cd-131">Descrição</span><span class="sxs-lookup"><span data-stu-id="330cd-131">Description</span></span> | <span data-ttu-id="330cd-132">Objetos com suporte</span><span class="sxs-lookup"><span data-stu-id="330cd-132">Supported objects</span></span> |
|:---------------|:-------------|:-----------|
| `onAdded` | <span data-ttu-id="330cd-133">Evento que ocorre quando um gráfico é adicionado.</span><span class="sxs-lookup"><span data-stu-id="330cd-133">Event that occurs when an object is added.</span></span> | [<span data-ttu-id="330cd-134">**ChartCollection**</span><span class="sxs-lookup"><span data-stu-id="330cd-134">**chartCollection**</span></span>](https://github.com/OfficeDev/office-js-docs/blob/ExcelJs_OpenSpec/reference/new-events.md) |
| `onDeleted` | <span data-ttu-id="330cd-135">Evento que ocorre quando um gráfico é excluído.</span><span class="sxs-lookup"><span data-stu-id="330cd-135">Event that occurs when an object is deleted.</span></span> | [<span data-ttu-id="330cd-136">**ChartCollection**</span><span class="sxs-lookup"><span data-stu-id="330cd-136">**chartCollection**</span></span>](https://github.com/OfficeDev/office-js-docs/blob/ExcelJs_OpenSpec/reference/new-events.md) |
| `onActivated` | <span data-ttu-id="330cd-137">Evento que ocorre quando um gráfico é ativado.</span><span class="sxs-lookup"><span data-stu-id="330cd-137">Event that occurs when an object is activated.</span></span> | <span data-ttu-id="330cd-138">[**Gráfico**](https://github.com/OfficeDev/office-js-docs/blob/ExcelJs_OpenSpec/reference/new-events.md), [**ChartCollection**](https://github.com/OfficeDev/office-js-docs/blob/ExcelJs_OpenSpec/reference/new-events.md)</span><span class="sxs-lookup"><span data-stu-id="330cd-138">[**Chart**](https://github.com/OfficeDev/office-js-docs/blob/ExcelJs_OpenSpec/reference/new-events.md), [**ChartCollection**](https://github.com/OfficeDev/office-js-docs/blob/ExcelJs_OpenSpec/reference/new-events.md)</span></span> |
| `onDeactivated` | <span data-ttu-id="330cd-139">Evento que ocorre quando um gráfico é desativado.</span><span class="sxs-lookup"><span data-stu-id="330cd-139">Event that occurs when an object is deactivated.</span></span> | <span data-ttu-id="330cd-140">[**Gráfico**](https://github.com/OfficeDev/office-js-docs/blob/ExcelJs_OpenSpec/reference/new-events.md), [**ChartCollection**](https://github.com/OfficeDev/office-js-docs/blob/ExcelJs_OpenSpec/reference/new-events.md)</span><span class="sxs-lookup"><span data-stu-id="330cd-140">[**Chart**](https://github.com/OfficeDev/office-js-docs/blob/ExcelJs_OpenSpec/reference/new-events.md), [**ChartCollection**](https://github.com/OfficeDev/office-js-docs/blob/ExcelJs_OpenSpec/reference/new-events.md)</span></span> |
| `onCalculated` | <span data-ttu-id="330cd-141">Evento que ocorre quando uma planilha termina o cálculo (ou todas as planilhas da coleção foram concluídas).</span><span class="sxs-lookup"><span data-stu-id="330cd-141">Event that occurs when a worksheet has finished calculation (or all the worksheets of the collection have finished).</span></span> | <span data-ttu-id="330cd-142">[**WorksheetCollection**](https://github.com/OfficeDev/office-js-docs/blob/ExcelJs_OpenSpec/reference/new-events.md), [**Planilha**](https://github.com/OfficeDev/office-js-docs/blob/ExcelJs_OpenSpec/reference/new-events.md)</span><span class="sxs-lookup"><span data-stu-id="330cd-142">**WorksheetCollection**, [Worksheet](https://github.com/OfficeDev/office-js-docs/blob/ExcelJs_OpenSpec/reference/new-events.md)</span></span> |

### <a name="event-triggers"></a><span data-ttu-id="330cd-143">Gatilhos de eventos</span><span class="sxs-lookup"><span data-stu-id="330cd-143">Event triggers</span></span>

<span data-ttu-id="330cd-144">Os eventos em uma pasta de trabalho do Excel podem ser acionados por:</span><span class="sxs-lookup"><span data-stu-id="330cd-144">Events within an Excel workbook can be triggered by:</span></span>

- <span data-ttu-id="330cd-145">Interação do usuário por meio da interface do usuário (UI) do Excel que altere a pasta de trabalho</span><span class="sxs-lookup"><span data-stu-id="330cd-145">User interaction via the Excel user interface (UI) that changes the workbook</span></span>
- <span data-ttu-id="330cd-146">Código de suplemento do Office (em JavaScript) que altere a pasta de trabalho</span><span class="sxs-lookup"><span data-stu-id="330cd-146">Office add-in (JavaScript) code that changes the workbook</span></span>
- <span data-ttu-id="330cd-147">Código de suplemento de VBA (macro) que altere a pasta de trabalho</span><span class="sxs-lookup"><span data-stu-id="330cd-147">VBA add-in (macro) code that changes the workbook</span></span>

<span data-ttu-id="330cd-148">Todas as alterações que sejam compatíveis com o comportamento padrão do Excel acionarão eventos correspondentes em uma pasta de trabalho.</span><span class="sxs-lookup"><span data-stu-id="330cd-148">Any change that complies with default behavior of Excel will trigger the corresponding event(s) in a workbook.</span></span>

### <a name="lifecycle-of-an-event-handler"></a><span data-ttu-id="330cd-149">Ciclo de vida de um manipulador de eventos</span><span class="sxs-lookup"><span data-stu-id="330cd-149">Lifecycle of an event handler</span></span>

<span data-ttu-id="330cd-p103">Um manipulador de eventos é criado quando um suplemento o registra e é destruído quando o suplemento cancela seu registro ou quando o suplemento for fechado. Os manipuladores de eventos não persistem como parte do arquivo de Excel.</span><span class="sxs-lookup"><span data-stu-id="330cd-p103">An event handler is created when an add-in registers the event handler and is destroyed when the add-in unregisters the event handler or when the add-in is closed. Event handlers do not persist as part of the Excel file.</span></span>

### <a name="events-and-coauthoring"></a><span data-ttu-id="330cd-152">Eventos e coautoria</span><span class="sxs-lookup"><span data-stu-id="330cd-152">Events and coauthoring</span></span>

<span data-ttu-id="330cd-p104">Com a [coautoria](co-authoring-in-excel-add-ins.md), várias pessoas podem trabalhar em conjunto e editar a mesma pasta de trabalho do Excel simultaneamente. Em eventos que podem ser disparados por um coautor, como `onChanged`, o objeto de **evento** respectivo conterá a propriedade **fonte** que indica se o evento foi acionado localmente pelo usuário atual (`event.source = Local`) ou pelo coautor remoto (`event.source = Remote`).</span><span class="sxs-lookup"><span data-stu-id="330cd-p104">With [coauthoring](co-authoring-in-excel-add-ins.md), multiple people can work together and edit the same Excel workbook simultaneously. For events that can be triggered by a coauthor, such as `onChanged`, the corresponding **Event** object will contain a **source** property that indicates whether the event was triggered locally by the current user (`event.source = Local`) or was triggered by the remote coauthor (`event.source = Remote`).</span></span>

## <a name="register-an-event-handler"></a><span data-ttu-id="330cd-155">Registrar um manipulador de eventos.</span><span class="sxs-lookup"><span data-stu-id="330cd-155">Register an event handler</span></span>

<span data-ttu-id="330cd-156">O exemplo de código a seguir registra um manipulador de eventos para o evento `onChanged` na planilha **Sample**.</span><span class="sxs-lookup"><span data-stu-id="330cd-156">The following code sample registers an event handler for the `onChanged` event in the worksheet named **Sample**.</span></span> <span data-ttu-id="330cd-157">O código especifica que, quando os dados forem alterados na planilha, a função `handleDataChange` deve ser executada.</span><span class="sxs-lookup"><span data-stu-id="330cd-157">The code specifies that when data changes in that worksheet, the `handleDataChange` function should run.</span></span>

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

## <a name="handle-an-event"></a><span data-ttu-id="330cd-158">Manipular um evento</span><span class="sxs-lookup"><span data-stu-id="330cd-158">Handle an event</span></span>

<span data-ttu-id="330cd-159">Como mostrado no exemplo anterior, quando você registrar um manipulador de eventos, indica a função a ser executada quando o evento especificado ocorre.</span><span class="sxs-lookup"><span data-stu-id="330cd-159">As shown in the previous example, when you register an event handler, you indicate the function that should run when the specified event occurs.</span></span> <span data-ttu-id="330cd-160">Você pode criar essa função para executar as ações que seu cenário exige.</span><span class="sxs-lookup"><span data-stu-id="330cd-160">You can design that function to perform whatever actions your scenario requires.</span></span> <span data-ttu-id="330cd-161">O exemplo de código a seguir mostra uma função de manipulador de eventos que simplesmente grava informações sobre o evento no console.</span><span class="sxs-lookup"><span data-stu-id="330cd-161">The following code sample shows an event handler function that simply writes information about the event to the console.</span></span> 

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

## <a name="remove-an-event-handler"></a><span data-ttu-id="330cd-162">Remover um manipulador de eventos</span><span class="sxs-lookup"><span data-stu-id="330cd-162">Remove an event handler</span></span>

<span data-ttu-id="330cd-163">O exemplo de código a seguir registra um manipulador de eventos para o evento `onSelectionChanged` na planilha **Sample** e define a função `handleSelectionChange` a executar quando o evento ocorrer.</span><span class="sxs-lookup"><span data-stu-id="330cd-163">The following code sample registers an event handler for the `onSelectionChanged` event in the worksheet named **Sample** and defines the `handleSelectionChange` function that will run when the event occurs.</span></span> <span data-ttu-id="330cd-164">Também define a função `remove()` que pode ser chamada posteriormente para remover aquele manipulador de eventos.</span><span class="sxs-lookup"><span data-stu-id="330cd-164">It also defines the `remove()` function that can subsequently be called to remove that event handler.</span></span>

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

## <a name="enable-and-disable-events"></a><span data-ttu-id="330cd-165">Ativar e desativar eventos</span><span class="sxs-lookup"><span data-stu-id="330cd-165">Enable and disable agents</span></span>

> [!NOTE]
> <span data-ttu-id="330cd-166">Este recurso está atualmente disponível somente na versão prévia pública (beta).</span><span class="sxs-lookup"><span data-stu-id="330cd-166">This feature is currently available only in public preview (beta).</span></span> <span data-ttu-id="330cd-167">Para usá-lo, você deve fazer referência a biblioteca beta da CDN do Office. js: https://appsforoffice.microsoft.com/lib/beta/hosted/office.js.</span><span class="sxs-lookup"><span data-stu-id="330cd-167">To use these features, you must use the beta library of the Office.js CDN: https://appsforoffice.microsoft.com/lib/beta/hosted/office.js.</span></span>

<span data-ttu-id="330cd-168">Eventos são ativados e desativados em [tempo de execução](https://docs.microsoft.com/en-us/javascript/api/excel/excel.runtime?view=office-js).</span><span class="sxs-lookup"><span data-stu-id="330cd-168">Events are turned on and off at the [runtime](https://docs.microsoft.com/en-us/javascript/api/excel/excel.runtime?view=office-js) level.</span></span> <span data-ttu-id="330cd-169">A propriedade `enableEvents` determina se os eventos são disparados e seus manipuladores serão ativados.</span><span class="sxs-lookup"><span data-stu-id="330cd-169">The `enableEvents` property determines if events are fired and their handlers are activated.</span></span> <span data-ttu-id="330cd-170">Desativar eventos é útil quando o desempenho é crítico ou durante a edição de várias entidades e se deseja evitar o acionamento de eventos até terminar.</span><span class="sxs-lookup"><span data-stu-id="330cd-170">Turning events off is useful when performance is critical or when editing multiple entities and want to avoid firing events until you have finished.</span></span>

<span data-ttu-id="330cd-171">O exemplo de código a seguir mostra como ativar e desativar eventos.</span><span class="sxs-lookup"><span data-stu-id="330cd-171">The following code sample shows how to toggle events on and off.</span></span>

```typescript
async function toggleEvents() {
    await Excel.run(async (context) => {
        context.runtime.load("enableEvents");
        await context.sync();
        const eventBoolean = !context.runtime.enableEvents
        context.runtime.enableEvents = eventBoolean;
        if (eventBoolean) {
            console.log("Events are currently on.");
        } else {
            console.log("Events are currently off.");
        }
        await context.sync();
    });
}
```

## <a name="see-also"></a><span data-ttu-id="330cd-172">Veja também</span><span class="sxs-lookup"><span data-stu-id="330cd-172">See also</span></span>

- [<span data-ttu-id="330cd-173">Principais conceitos da API JavaScript do Excel</span><span class="sxs-lookup"><span data-stu-id="330cd-173">Excel JavaScript API core concepts</span></span>](excel-add-ins-core-concepts.md)
- [<span data-ttu-id="330cd-174">Especificação para abrir API JavaScript do Excel</span><span class="sxs-lookup"><span data-stu-id="330cd-174">Excel JavaScript API Open Specification</span></span>](https://github.com/OfficeDev/office-js-docs/tree/ExcelJs_OpenSpec)