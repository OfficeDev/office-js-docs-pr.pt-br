---
title: Trabalhar com eventos usando a API JavaScript do Excel
description: ''
ms.date: 05/25/2018
ms.openlocfilehash: b928910cc673cfe8ff99906259b51fa2c3afdca4
ms.sourcegitcommit: 17f60431644b448a4816913039aaebfa328f9b0a
ms.translationtype: HT
ms.contentlocale: pt-BR
ms.lasthandoff: 05/25/2018
---
# <a name="work-with-events-using-the-excel-javascript-api"></a><span data-ttu-id="4b5cb-102">Trabalhar com eventos usando a API JavaScript do Excel</span><span class="sxs-lookup"><span data-stu-id="4b5cb-102">Work with Events using the Excel JavaScript API</span></span> 

<span data-ttu-id="4b5cb-103">Este artigo descreve conceitos importantes relacionados ao trabalho com eventos no Excel e fornece exemplos de c?digo que mostram como registrar manipuladores de eventos, lidar com eventos e remover manipuladores de eventos usando as APIs JavaScript do Excel.</span><span class="sxs-lookup"><span data-stu-id="4b5cb-103">This article describes important concepts related to working with events in Excel and provides code samples that show how to register event handlers, handle events, and remove event handlers using the Excel JavaScript API.</span></span> 

## <a name="events-in-excel"></a><span data-ttu-id="4b5cb-104">Eventos no Excel</span><span class="sxs-lookup"><span data-stu-id="4b5cb-104">Events in Excel</span></span>

<span data-ttu-id="4b5cb-105">Sempre que ocorrerem certos tipos de altera??es em uma pasta de trabalho do Excel, uma notifica??o do evento ser? ativada.</span><span class="sxs-lookup"><span data-stu-id="4b5cb-105">Each time certain types of changes occur in an Excel workbook, an event notification fires.</span></span> <span data-ttu-id="4b5cb-106">Ao usar as APIs JavaScript do Excel, voc? pode registrar manipuladores de eventos que permitem que o suplemento execute automaticamente uma fun??o designada quando ocorre um evento espec?fico.</span><span class="sxs-lookup"><span data-stu-id="4b5cb-106">By using the Excel JavaScript API, you can register event handlers that allow your add-in to automatically run a designated function when a specific event occurs.</span></span> <span data-ttu-id="4b5cb-107">Os eventos a seguir t?m suporte no momento:</span><span class="sxs-lookup"><span data-stu-id="4b5cb-107">The following events are currently supported.</span></span>

| <span data-ttu-id="4b5cb-108">Evento</span><span class="sxs-lookup"><span data-stu-id="4b5cb-108">Event</span></span> | <span data-ttu-id="4b5cb-109">Descri??o</span><span class="sxs-lookup"><span data-stu-id="4b5cb-109">Description</span></span> | <span data-ttu-id="4b5cb-110">Objetos com suporte</span><span class="sxs-lookup"><span data-stu-id="4b5cb-110">Supported objects</span></span> |
|:---------------|:-------------|:-----------|
| `onAdded` | <span data-ttu-id="4b5cb-111">Evento que ocorre quando um objeto ? adicionado.</span><span class="sxs-lookup"><span data-stu-id="4b5cb-111">Event that occurs when an object is added.</span></span> | [<span data-ttu-id="4b5cb-112">**WorksheetCollection**</span><span class="sxs-lookup"><span data-stu-id="4b5cb-112">**WorksheetCollection**</span></span>](https://dev.office.com/reference/add-ins/excel/worksheetcollection) |
| `onDeleted` | <span data-ttu-id="4b5cb-113">Evento que ocorre quando um objeto ? exclu?do.</span><span class="sxs-lookup"><span data-stu-id="4b5cb-113">Event that occurs when an object is deleted.</span></span> | [<span data-ttu-id="4b5cb-114">**WorksheetCollection**</span><span class="sxs-lookup"><span data-stu-id="4b5cb-114">**WorksheetCollection**</span></span>](https://dev.office.com/reference/add-ins/excel/worksheetcollection) |
| `onActivated` | <span data-ttu-id="4b5cb-115">Evento que ocorre quando um objeto ? ativado.</span><span class="sxs-lookup"><span data-stu-id="4b5cb-115">Event that occurs when an object is activated.</span></span> | <span data-ttu-id="4b5cb-116">[**WorksheetCollection**](https://dev.office.com/reference/add-ins/excel/worksheetcollection), [**Planilha**](https://dev.office.com/reference/add-ins/excel/worksheet)</span><span class="sxs-lookup"><span data-stu-id="4b5cb-116">**WorksheetCollection**, [Worksheet](https://dev.office.com/reference/add-ins/excel/worksheetcollection)</span></span> |
| `onDeactivated` | <span data-ttu-id="4b5cb-117">Evento que ocorre quando um objeto ? desativado.</span><span class="sxs-lookup"><span data-stu-id="4b5cb-117">Event that occurs when an object is deactivated.</span></span> | <span data-ttu-id="4b5cb-118">[**WorksheetCollection**](https://dev.office.com/reference/add-ins/excel/worksheetcollection), [**Planilha**](https://dev.office.com/reference/add-ins/excel/worksheet)</span><span class="sxs-lookup"><span data-stu-id="4b5cb-118">**WorksheetCollection**, [Worksheet](https://dev.office.com/reference/add-ins/excel/worksheetcollection)</span></span> |
| `onChanged` | <span data-ttu-id="4b5cb-119">Evento que ocorre quando os dados das c?lulas s?o alterados.</span><span class="sxs-lookup"><span data-stu-id="4b5cb-119">Event that occurs when data within cells is changed.</span></span> | <span data-ttu-id="4b5cb-120">[**Planilha**](https://dev.office.com/reference/add-ins/excel/worksheet), [**Tabela**](https://dev.office.com/reference/add-ins/excel/table), [**TableCollection**](https://dev.office.com/reference/add-ins/excel/tablecollection)</span><span class="sxs-lookup"><span data-stu-id="4b5cb-120">[**Worksheet**](https://dev.office.com/reference/add-ins/excel/worksheet), [**Table**](https://dev.office.com/reference/add-ins/excel/table), [**TableCollection**](https://dev.office.com/reference/add-ins/excel/tablecollection)</span></span> |
| `onDataChanged` | <span data-ttu-id="4b5cb-121">Evento que ocorre quando os dados ou a formata??o dentro da associa??o s?o alterados.</span><span class="sxs-lookup"><span data-stu-id="4b5cb-121">Occurs when data or formatting within the binding is changed.</span></span> | [<span data-ttu-id="4b5cb-122">**Associa??o**</span><span class="sxs-lookup"><span data-stu-id="4b5cb-122">**Binding**</span></span>](https://dev.office.com/reference/add-ins/excel/binding) |
| `onSelectionChanged` | <span data-ttu-id="4b5cb-123">Evento que ocorre quando uma c?lula ativa ou um intervalo selecionado s?o alterados.</span><span class="sxs-lookup"><span data-stu-id="4b5cb-123">Event that occurs when the active cell or selected range is changed.</span></span> | <span data-ttu-id="4b5cb-124">[**Planilha**](https://dev.office.com/reference/add-ins/excel/worksheet), [**Tabela**](https://dev.office.com/reference/add-ins/excel/table), [**Associa??o**](https://dev.office.com/reference/add-ins/excel/binding)</span><span class="sxs-lookup"><span data-stu-id="4b5cb-124">**Worksheet**, [Table](https://dev.office.com/reference/add-ins/excel/worksheet), **Binding**</span></span> |

### <a name="event-triggers"></a><span data-ttu-id="4b5cb-125">Gatilhos de eventos</span><span class="sxs-lookup"><span data-stu-id="4b5cb-125">Event triggers</span></span>

<span data-ttu-id="4b5cb-126">Os eventos em uma pasta de trabalho do Excel podem ser acionados por:</span><span class="sxs-lookup"><span data-stu-id="4b5cb-126">Events within an Excel workbook can be triggered by:</span></span>

- <span data-ttu-id="4b5cb-127">Intera??o do usu?rio por meio da interface do usu?rio (UI) do Excel que altere a pasta de trabalho</span><span class="sxs-lookup"><span data-stu-id="4b5cb-127">User interaction via the Excel user interface (UI) that changes the workbook</span></span>
- <span data-ttu-id="4b5cb-128">C?digo de suplemento do Office (em JavaScript) que altere a pasta de trabalho</span><span class="sxs-lookup"><span data-stu-id="4b5cb-128">Office add-in (JavaScript) code that changes the workbook</span></span>
- <span data-ttu-id="4b5cb-129">C?digo de suplemento de VBA (macro) que altere a pasta de trabalho</span><span class="sxs-lookup"><span data-stu-id="4b5cb-129">VBA add-in (macro) code that changes the workbook</span></span>

<span data-ttu-id="4b5cb-130">Todas as altera??es que sejam compat?veis com o comportamento padr?o do Excel acionar?o eventos correspondentes em uma pasta de trabalho.</span><span class="sxs-lookup"><span data-stu-id="4b5cb-130">Any change that complies with default behavior of Excel will trigger the corresponding event(s) in a workbook.</span></span>

### <a name="lifecycle-of-an-event-handler"></a><span data-ttu-id="4b5cb-131">Ciclo de vida de um manipulador de eventos</span><span class="sxs-lookup"><span data-stu-id="4b5cb-131">Lifecycle of an event handler</span></span>

<span data-ttu-id="4b5cb-p102">Um manipulador de eventos ? criado quando um suplemento o registra e ? destru?do quando o suplemento cancela seu registro ou quando o suplemento for fechado. Os manipuladores de eventos n?o persistem como parte do arquivo de Excel.</span><span class="sxs-lookup"><span data-stu-id="4b5cb-p102">An event handler is created when an add-in registers the event handler and is destroyed when the add-in unregisters the event handler or when the add-in is closed. Event handlers do not persist as part of the Excel file.</span></span>

### <a name="events-and-coauthoring"></a><span data-ttu-id="4b5cb-134">Eventos e coautoria</span><span class="sxs-lookup"><span data-stu-id="4b5cb-134">Events and coauthoring</span></span>

<span data-ttu-id="4b5cb-p103">Com a [coautoria](co-authoring-in-excel-add-ins.md), v?rias pessoas podem trabalhar em conjunto e editar a mesma pasta de trabalho do Excel simultaneamente. Em eventos que podem ser disparados por um coautor, como `onChanged`, o objeto de **evento** respectivo conter? a propriedade **fonte** que indica se o evento foi acionado localmente pelo usu?rio atual (`event.source = Local`) ou pelo coautor remoto (`event.source = Remote`).</span><span class="sxs-lookup"><span data-stu-id="4b5cb-p103">With [coauthoring](co-authoring-in-excel-add-ins.md), multiple people can work together and edit the same Excel workbook simultaneously. For events that can be triggered by a coauthor, such as `onChanged`, the corresponding **Event** object will contain a **source** property that indicates whether the event was triggered locally by the current user (`event.source = Local`) or was triggered by the remote coauthor (`event.source = Remote`).</span></span>

## <a name="register-an-event-handler"></a><span data-ttu-id="4b5cb-137">Registrar um manipulador de eventos.</span><span class="sxs-lookup"><span data-stu-id="4b5cb-137">Register an event handler</span></span>

<span data-ttu-id="4b5cb-138">O exemplo de c?digo a seguir registra um manipulador de eventos para o evento `onChanged` na planilha **Sample**.</span><span class="sxs-lookup"><span data-stu-id="4b5cb-138">The following code sample registers an event handler for the `onChanged` event in the worksheet named **Sample**.</span></span> <span data-ttu-id="4b5cb-139">O c?digo especifica que, quando os dados forem alterados na planilha, a fun??o `handleDataChange` deve ser executada.</span><span class="sxs-lookup"><span data-stu-id="4b5cb-139">The code specifies that when data changes in that worksheet, the `handleDataChange` function should run.</span></span>

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

## <a name="handle-an-event"></a><span data-ttu-id="4b5cb-140">Manipular um evento</span><span class="sxs-lookup"><span data-stu-id="4b5cb-140">Handle an event</span></span>

<span data-ttu-id="4b5cb-141">Como mostrado no exemplo anterior, quando voc? registrar um manipulador de eventos, indica a fun??o a ser executada quando o evento especificado ocorre.</span><span class="sxs-lookup"><span data-stu-id="4b5cb-141">As shown in the previous example, when you register an event handler, you indicate the function that should run when the specified event occurs.</span></span> <span data-ttu-id="4b5cb-142">Voc? pode criar essa fun??o para executar as a??es que seu cen?rio exige.</span><span class="sxs-lookup"><span data-stu-id="4b5cb-142">You can design that function to perform whatever actions your scenario requires.</span></span> <span data-ttu-id="4b5cb-143">O exemplo de c?digo a seguir mostra uma fun??o de manipulador de eventos que simplesmente grava informa??es sobre o evento no console.</span><span class="sxs-lookup"><span data-stu-id="4b5cb-143">The following code sample shows an event handler function that simply writes information about the event to the console.</span></span> 

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

## <a name="remove-an-event-handler"></a><span data-ttu-id="4b5cb-144">Remover um manipulador de eventos</span><span class="sxs-lookup"><span data-stu-id="4b5cb-144">Remove an event handler</span></span>

<span data-ttu-id="4b5cb-145">O exemplo de c?digo a seguir registra um manipulador de eventos para o evento `onSelectionChanged` na planilha **Sample** e define a fun??o `handleSelectionChange` a executar quando o evento ocorrer.</span><span class="sxs-lookup"><span data-stu-id="4b5cb-145">The following code sample registers an event handler for the `onSelectionChanged` event in the worksheet named **Sample** and defines the `handleSelectionChange` function that will run when the event occurs.</span></span> <span data-ttu-id="4b5cb-146">Tamb?m define a fun??o `remove()` que pode ser chamada posteriormente para remover aquele manipulador de eventos.</span><span class="sxs-lookup"><span data-stu-id="4b5cb-146">It also defines the `remove()` function that can subsequently be called to remove that event handler.</span></span>

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

## <a name="see-also"></a><span data-ttu-id="4b5cb-147">Confira tamb?m</span><span class="sxs-lookup"><span data-stu-id="4b5cb-147">See also</span></span>

- [<span data-ttu-id="4b5cb-148">Principais conceitos da API JavaScript do Excel</span><span class="sxs-lookup"><span data-stu-id="4b5cb-148">Excel JavaScript API core concepts</span></span>](excel-add-ins-core-concepts.md)
- [<span data-ttu-id="4b5cb-149">Especifica??o para abrir API JavaScript do Excel</span><span class="sxs-lookup"><span data-stu-id="4b5cb-149">Excel JavaScript API Open Specification</span></span>](https://github.com/OfficeDev/office-js-docs/tree/ExcelJs_OpenSpec)