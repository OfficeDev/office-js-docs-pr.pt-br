---
title: Trabalhar com eventos usando a API JavaScript do Excel
description: ''
ms.date: 05/25/2018
ms.openlocfilehash: 575e4112ed5f55356020eed8327d309fc58cd643
ms.sourcegitcommit: 9685fd83136bd2106f4c5595bda0010bc1b1950b
ms.translationtype: HT
ms.contentlocale: pt-BR
ms.lasthandoff: 07/19/2018
ms.locfileid: "20596515"
---
# <a name="work-with-events-using-the-excel-javascript-api"></a><span data-ttu-id="e555a-102">Trabalhar com eventos usando a API JavaScript do Excel</span><span class="sxs-lookup"><span data-stu-id="e555a-102">Work with Events using the Excel JavaScript API</span></span> 

<span data-ttu-id="e555a-103">Este artigo descreve conceitos importantes relacionados ao trabalho com eventos no Excel e fornece exemplos de código que mostram como registrar manipuladores de eventos, lidar com eventos e remover manipuladores de eventos usando as APIs JavaScript do Excel.</span><span class="sxs-lookup"><span data-stu-id="e555a-103">This article describes important concepts related to working with events in Excel and provides code samples that show how to register event handlers, handle events, and remove event handlers using the Excel JavaScript API.</span></span> 

## <a name="events-in-excel"></a><span data-ttu-id="e555a-104">Eventos no Excel</span><span class="sxs-lookup"><span data-stu-id="e555a-104">Events in Excel</span></span>

<span data-ttu-id="e555a-105">Sempre que ocorrerem certos tipos de alterações em uma pasta de trabalho do Excel, uma notificação do evento será ativada.</span><span class="sxs-lookup"><span data-stu-id="e555a-105">Each time certain types of changes occur in an Excel workbook, an event notification fires.</span></span> <span data-ttu-id="e555a-106">Ao usar as APIs JavaScript do Excel, você pode registrar manipuladores de eventos que permitem que o suplemento execute automaticamente uma função designada quando ocorre um evento específico.</span><span class="sxs-lookup"><span data-stu-id="e555a-106">By using the Excel JavaScript API, you can register event handlers that allow your add-in to automatically run a designated function when a specific event occurs.</span></span> <span data-ttu-id="e555a-107">Os eventos a seguir têm suporte no momento:</span><span class="sxs-lookup"><span data-stu-id="e555a-107">The following events are currently supported.</span></span>

| <span data-ttu-id="e555a-108">Evento</span><span class="sxs-lookup"><span data-stu-id="e555a-108">Event</span></span> | <span data-ttu-id="e555a-109">Descrição</span><span class="sxs-lookup"><span data-stu-id="e555a-109">Description</span></span> | <span data-ttu-id="e555a-110">Objetos com suporte</span><span class="sxs-lookup"><span data-stu-id="e555a-110">Supported objects</span></span> |
|:---------------|:-------------|:-----------|
| `onAdded` | <span data-ttu-id="e555a-111">Evento que ocorre quando um objeto é adicionado.</span><span class="sxs-lookup"><span data-stu-id="e555a-111">Event that occurs when an object is added.</span></span> | [<span data-ttu-id="e555a-112">**WorksheetCollection**</span><span class="sxs-lookup"><span data-stu-id="e555a-112">**WorksheetCollection**</span></span>](https://dev.office.com/reference/add-ins/excel/worksheetcollection) |
| `onDeleted` | <span data-ttu-id="e555a-113">Evento que ocorre quando um objeto é excluído.</span><span class="sxs-lookup"><span data-stu-id="e555a-113">Event that occurs when an object is deleted.</span></span> | [<span data-ttu-id="e555a-114">**WorksheetCollection**</span><span class="sxs-lookup"><span data-stu-id="e555a-114">**WorksheetCollection**</span></span>](https://dev.office.com/reference/add-ins/excel/worksheetcollection) |
| `onActivated` | <span data-ttu-id="e555a-115">Evento que ocorre quando um objeto é ativado.</span><span class="sxs-lookup"><span data-stu-id="e555a-115">Event that occurs when an object is activated.</span></span> | <span data-ttu-id="e555a-116">[**WorksheetCollection**](https://dev.office.com/reference/add-ins/excel/worksheetcollection), [**Planilha**](https://dev.office.com/reference/add-ins/excel/worksheet)</span><span class="sxs-lookup"><span data-stu-id="e555a-116">**WorksheetCollection**, [Worksheet](https://dev.office.com/reference/add-ins/excel/worksheetcollection)</span></span> |
| `onDeactivated` | <span data-ttu-id="e555a-117">Evento que ocorre quando um objeto é desativado.</span><span class="sxs-lookup"><span data-stu-id="e555a-117">Event that occurs when an object is deactivated.</span></span> | <span data-ttu-id="e555a-118">[**WorksheetCollection**](https://dev.office.com/reference/add-ins/excel/worksheetcollection), [**Planilha**](https://dev.office.com/reference/add-ins/excel/worksheet)</span><span class="sxs-lookup"><span data-stu-id="e555a-118">**WorksheetCollection**, [Worksheet](https://dev.office.com/reference/add-ins/excel/worksheetcollection)</span></span> |
| `onChanged` | <span data-ttu-id="e555a-119">Evento que ocorre quando os dados de células são alterados.</span><span class="sxs-lookup"><span data-stu-id="e555a-119">Event that occurs when data within cells is changed.</span></span> | <span data-ttu-id="e555a-120">[**Planilha**](https://dev.office.com/reference/add-ins/excel/worksheet), [**Tabela**](https://dev.office.com/reference/add-ins/excel/table), [**TableCollection**](https://dev.office.com/reference/add-ins/excel/tablecollection)</span><span class="sxs-lookup"><span data-stu-id="e555a-120">[**Worksheet**](https://dev.office.com/reference/add-ins/excel/worksheet), [**Table**](https://dev.office.com/reference/add-ins/excel/table), [**TableCollection**](https://dev.office.com/reference/add-ins/excel/tablecollection)</span></span> |
| `onDataChanged` | <span data-ttu-id="e555a-121">Evento que ocorre quando os dados ou a formatação dentro da associação são alterados.</span><span class="sxs-lookup"><span data-stu-id="e555a-121">Occurs when data or formatting within the binding is changed.</span></span> | [<span data-ttu-id="e555a-122">**Associação**</span><span class="sxs-lookup"><span data-stu-id="e555a-122">**Binding**</span></span>](https://dev.office.com/reference/add-ins/excel/binding) |
| `onSelectionChanged` | <span data-ttu-id="e555a-123">Evento que ocorre quando uma célula ativa ou um intervalo selecionado são alterados.</span><span class="sxs-lookup"><span data-stu-id="e555a-123">Event that occurs when the active cell or selected range is changed.</span></span> | <span data-ttu-id="e555a-124">[**Planilha**](https://dev.office.com/reference/add-ins/excel/worksheet), [**Tabela**](https://dev.office.com/reference/add-ins/excel/table), [**Associação**](https://dev.office.com/reference/add-ins/excel/binding)</span><span class="sxs-lookup"><span data-stu-id="e555a-124">**Worksheet**, [Table](https://dev.office.com/reference/add-ins/excel/worksheet), **Binding**</span></span> |
| `onSettingsChanged` | <span data-ttu-id="e555a-125">Evento que ocorre quando as Configurações no documento são alteradas.</span><span class="sxs-lookup"><span data-stu-id="e555a-125">Occurs when the Settings in the document are changed.</span></span> | [<span data-ttu-id="e555a-126">**SettingCollection**</span><span class="sxs-lookup"><span data-stu-id="e555a-126">**SettingCollection**</span></span>](https://dev.office.com/reference/add-ins/excel/settingcollection) |

### <a name="event-triggers"></a><span data-ttu-id="e555a-127">Gatilhos de eventos</span><span class="sxs-lookup"><span data-stu-id="e555a-127">Event triggers</span></span>

<span data-ttu-id="e555a-128">Os eventos em uma pasta de trabalho do Excel podem ser acionados por:</span><span class="sxs-lookup"><span data-stu-id="e555a-128">Events within an Excel workbook can be triggered by:</span></span>

- <span data-ttu-id="e555a-129">Interação do usuário por meio da interface do usuário (UI) do Excel que altere a pasta de trabalho</span><span class="sxs-lookup"><span data-stu-id="e555a-129">User interaction via the Excel user interface (UI) that changes the workbook</span></span>
- <span data-ttu-id="e555a-130">Código de suplemento do Office (em JavaScript) que altere a pasta de trabalho</span><span class="sxs-lookup"><span data-stu-id="e555a-130">Office add-in (JavaScript) code that changes the workbook</span></span>
- <span data-ttu-id="e555a-131">Código de suplemento de VBA (macro) que altere a pasta de trabalho</span><span class="sxs-lookup"><span data-stu-id="e555a-131">VBA add-in (macro) code that changes the workbook</span></span>

<span data-ttu-id="e555a-132">Todas as alterações que sejam compatíveis com o comportamento padrão do Excel acionarão eventos correspondentes em uma pasta de trabalho.</span><span class="sxs-lookup"><span data-stu-id="e555a-132">Any change that complies with default behavior of Excel will trigger the corresponding event(s) in a workbook.</span></span>

### <a name="lifecycle-of-an-event-handler"></a><span data-ttu-id="e555a-133">Ciclo de vida de um manipulador de eventos</span><span class="sxs-lookup"><span data-stu-id="e555a-133">Lifecycle of an event handler</span></span>

<span data-ttu-id="e555a-p102">Um manipulador de eventos é criado quando um suplemento o registra e é destruído quando o suplemento cancela seu registro ou quando o suplemento for fechado. Os manipuladores de eventos não persistem como parte do arquivo de Excel.</span><span class="sxs-lookup"><span data-stu-id="e555a-p102">An event handler is created when an add-in registers the event handler and is destroyed when the add-in unregisters the event handler or when the add-in is closed. Event handlers do not persist as part of the Excel file.</span></span>

### <a name="events-and-coauthoring"></a><span data-ttu-id="e555a-136">Eventos e coautoria</span><span class="sxs-lookup"><span data-stu-id="e555a-136">Events and coauthoring</span></span>

<span data-ttu-id="e555a-p103">Com a [coautoria](co-authoring-in-excel-add-ins.md), várias pessoas podem trabalhar em conjunto e editar a mesma pasta de trabalho do Excel simultaneamente. Em eventos que podem ser disparados por um coautor, como `onChanged`, o objeto de **evento** respectivo conterá a propriedade **fonte** que indica se o evento foi acionado localmente pelo usuário atual (`event.source = Local`) ou pelo coautor remoto (`event.source = Remote`).</span><span class="sxs-lookup"><span data-stu-id="e555a-p103">With [coauthoring](co-authoring-in-excel-add-ins.md), multiple people can work together and edit the same Excel workbook simultaneously. For events that can be triggered by a coauthor, such as `onChanged`, the corresponding **Event** object will contain a **source** property that indicates whether the event was triggered locally by the current user (`event.source = Local`) or was triggered by the remote coauthor (`event.source = Remote`).</span></span>

## <a name="register-an-event-handler"></a><span data-ttu-id="e555a-139">Registrar um manipulador de eventos.</span><span class="sxs-lookup"><span data-stu-id="e555a-139">Register an event handler</span></span>

<span data-ttu-id="e555a-140">O exemplo de código a seguir registra um manipulador de eventos para o evento `onChanged` na planilha **Sample**.</span><span class="sxs-lookup"><span data-stu-id="e555a-140">The following code sample registers an event handler for the `onChanged` event in the worksheet named **Sample**.</span></span> <span data-ttu-id="e555a-141">O código especifica que, quando os dados forem alterados na planilha, a função `handleDataChange` deve ser executada.</span><span class="sxs-lookup"><span data-stu-id="e555a-141">The code specifies that when data changes in that worksheet, the `handleDataChange` function should run.</span></span>

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

## <a name="handle-an-event"></a><span data-ttu-id="e555a-142">Manipular um evento</span><span class="sxs-lookup"><span data-stu-id="e555a-142">Handle an event</span></span>

<span data-ttu-id="e555a-143">Como mostrado no exemplo anterior, quando você registrar um manipulador de eventos, indica a função a ser executada quando o evento especificado ocorre.</span><span class="sxs-lookup"><span data-stu-id="e555a-143">As shown in the previous example, when you register an event handler, you indicate the function that should run when the specified event occurs.</span></span> <span data-ttu-id="e555a-144">Você pode criar essa função para executar as ações que seu cenário exige.</span><span class="sxs-lookup"><span data-stu-id="e555a-144">You can design that function to perform whatever actions your scenario requires.</span></span> <span data-ttu-id="e555a-145">O exemplo de código a seguir mostra uma função de manipulador de eventos que simplesmente grava informações sobre o evento no console.</span><span class="sxs-lookup"><span data-stu-id="e555a-145">The following code sample shows an event handler function that simply writes information about the event to the console.</span></span> 

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

## <a name="remove-an-event-handler"></a><span data-ttu-id="e555a-146">Remover um manipulador de eventos</span><span class="sxs-lookup"><span data-stu-id="e555a-146">Remove an event handler</span></span>

<span data-ttu-id="e555a-147">O exemplo de código a seguir registra um manipulador de eventos para o evento `onSelectionChanged` na planilha **Sample** e define a função `handleSelectionChange` a executar quando o evento ocorrer.</span><span class="sxs-lookup"><span data-stu-id="e555a-147">The following code sample registers an event handler for the `onSelectionChanged` event in the worksheet named **Sample** and defines the `handleSelectionChange` function that will run when the event occurs.</span></span> <span data-ttu-id="e555a-148">Também define a função `remove()` que pode ser chamada posteriormente para remover aquele manipulador de eventos.</span><span class="sxs-lookup"><span data-stu-id="e555a-148">It also defines the `remove()` function that can subsequently be called to remove that event handler.</span></span>

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

## <a name="see-also"></a><span data-ttu-id="e555a-149">Confira também</span><span class="sxs-lookup"><span data-stu-id="e555a-149">See also</span></span>

- [<span data-ttu-id="e555a-150">Principais conceitos da API JavaScript do Excel</span><span class="sxs-lookup"><span data-stu-id="e555a-150">Excel JavaScript API core concepts</span></span>](excel-add-ins-core-concepts.md)
- [<span data-ttu-id="e555a-151">Especificação para abrir API JavaScript do Excel</span><span class="sxs-lookup"><span data-stu-id="e555a-151">Excel JavaScript API Open Specification</span></span>](https://github.com/OfficeDev/office-js-docs/tree/ExcelJs_OpenSpec)