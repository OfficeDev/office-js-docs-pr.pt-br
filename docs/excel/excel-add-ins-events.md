---
title: Trabalhar com eventos usando a API JavaScript do Excel
description: ''
ms.date: 10/17/2018
ms.openlocfilehash: c3fbdf27dcbedf0d006973e6ebc2e01b02e6cec2
ms.sourcegitcommit: a6d6348075c1abed76d2146ddfc099b0151fe403
ms.translationtype: HT
ms.contentlocale: pt-BR
ms.lasthandoff: 10/19/2018
ms.locfileid: "25639935"
---
# <a name="work-with-events-using-the-excel-javascript-api"></a><span data-ttu-id="a5a55-102">Trabalhar com eventos usando a API JavaScript do Excel</span><span class="sxs-lookup"><span data-stu-id="a5a55-102">Work with Events using the Excel JavaScript API</span></span> 

<span data-ttu-id="a5a55-103">Este artigo descreve conceitos importantes relacionados ao trabalho com eventos no Excel e fornece exemplos de código que mostram como registrar manipuladores de eventos, lidar com eventos e remover manipuladores de eventos usando as APIs JavaScript do Excel.</span><span class="sxs-lookup"><span data-stu-id="a5a55-103">This article describes important concepts related to working with events in Excel and provides code samples that show how to register event handlers, handle events, and remove event handlers using the Excel JavaScript API.</span></span> 

## <a name="events-in-excel"></a><span data-ttu-id="a5a55-104">Eventos no Excel</span><span class="sxs-lookup"><span data-stu-id="a5a55-104">Events in Excel</span></span>

<span data-ttu-id="a5a55-p101">Sempre que ocorrerem determinados tipos de alterações em uma pasta de trabalho do Excel, uma notificação de evento é acionada. Usando a API JavaScript do Excel, você pode registrar manipuladores de eventos que permitem o suplemento executar automaticamente uma função designada, quando ocorre um evento específico. Os eventos a seguir são suportados no momento.</span><span class="sxs-lookup"><span data-stu-id="a5a55-p101">Each time certain types of changes occur in an Excel workbook, an event notification fires. By using the Excel JavaScript API, you can register event handlers that allow your add-in to automatically run a designated function when a specific event occurs. The following events are currently supported.</span></span>

| <span data-ttu-id="a5a55-108">Evento</span><span class="sxs-lookup"><span data-stu-id="a5a55-108">Event</span></span> | <span data-ttu-id="a5a55-109">Descrição</span><span class="sxs-lookup"><span data-stu-id="a5a55-109">Description</span></span> | <span data-ttu-id="a5a55-110">Objetos com suporte</span><span class="sxs-lookup"><span data-stu-id="a5a55-110">Supported objects</span></span> |
|:---------------|:-------------|:-----------|
| `onAdded` | <span data-ttu-id="a5a55-111">Evento que ocorre quando um objeto é adicionado.</span><span class="sxs-lookup"><span data-stu-id="a5a55-111">Event that occurs when an object is added.</span></span> | <span data-ttu-id="a5a55-112">[**ChartCollection**](https://docs.microsoft.com/javascript/api/excel/excel.chartcollection), [**WorksheetCollection**](https://docs.microsoft.com/javascript/api/excel/excel.worksheetcollection)</span><span class="sxs-lookup"><span data-stu-id="a5a55-112">[**ChartCollection**](https://docs.microsoft.com/javascript/api/excel/excel.chartcollection), [**WorksheetCollection**](https://docs.microsoft.com/javascript/api/excel/excel.worksheetcollection)</span></span> |
| `onDeleted` | <span data-ttu-id="a5a55-113">Evento que ocorre quando um objeto é excluído.</span><span class="sxs-lookup"><span data-stu-id="a5a55-113">Event that occurs when an object is deleted.</span></span> | <span data-ttu-id="a5a55-114">[**ChartCollection**](https://docs.microsoft.com/javascript/api/excel/excel.chartcollection), [**WorksheetCollection**](https://docs.microsoft.com/javascript/api/excel/excel.worksheetcollection)</span><span class="sxs-lookup"><span data-stu-id="a5a55-114">[**ChartCollection**](https://docs.microsoft.com/javascript/api/excel/excel.chartcollection), [**WorksheetCollection**](https://docs.microsoft.com/javascript/api/excel/excel.worksheetcollection)</span></span> |
| `onActivated` | <span data-ttu-id="a5a55-115">Evento que ocorre quando um objeto é ativado.</span><span class="sxs-lookup"><span data-stu-id="a5a55-115">Event that occurs when an object is activated.</span></span> | <span data-ttu-id="a5a55-116">[**Gráfico**](https://docs.microsoft.com/javascript/api/excel/excel.chart), [**ChartCollection**](https://docs.microsoft.com/javascript/api/excel/excel.chartcollection), [**WorksheetCollection**](https://docs.microsoft.com/javascript/api/excel/excel.worksheetcollection), [**Planilha**](https://docs.microsoft.com/javascript/api/excel/excel.worksheet)</span><span class="sxs-lookup"><span data-stu-id="a5a55-116">[**Chart**](https://docs.microsoft.com/javascript/api/excel/excel.chart), [**ChartCollection**](https://docs.microsoft.com/javascript/api/excel/excel.chartcollection), [**WorksheetCollection**](https://docs.microsoft.com/javascript/api/excel/excel.worksheetcollection), [**Worksheet**](https://docs.microsoft.com/javascript/api/excel/excel.worksheet)</span></span> |
| `onDeactivated` | <span data-ttu-id="a5a55-117">Evento que ocorre quando um objeto é desativado.</span><span class="sxs-lookup"><span data-stu-id="a5a55-117">Event that occurs when an object is deactivated.</span></span> | <span data-ttu-id="a5a55-118">[**Gráfico**](https://docs.microsoft.com/javascript/api/excel/excel.chart), [**ChartCollection**](https://docs.microsoft.com/javascript/api/excel/excel.chartcollection), [**WorksheetCollection**](https://docs.microsoft.com/javascript/api/excel/excel.worksheetcollection), [**Planilha**](https://docs.microsoft.com/javascript/api/excel/excel.worksheet)</span><span class="sxs-lookup"><span data-stu-id="a5a55-118">[**Chart**](https://docs.microsoft.com/javascript/api/excel/excel.chart), [**ChartCollection**](https://docs.microsoft.com/javascript/api/excel/excel.chartcollection), [**WorksheetCollection**](https://docs.microsoft.com/javascript/api/excel/excel.worksheetcollection), [**Worksheet**](https://docs.microsoft.com/javascript/api/excel/excel.worksheet)</span></span> |
| `onCalculated` | <span data-ttu-id="a5a55-119">Evento que ocorre quando uma planilha termina o cálculo (ou todas as planilhas da coleção foram concluídas).</span><span class="sxs-lookup"><span data-stu-id="a5a55-119">Event that occurs when a worksheet has finished calculation (or all the worksheets of the collection have finished).</span></span> | <span data-ttu-id="a5a55-120">[**WorksheetCollection**](https://docs.microsoft.com/javascript/api/excel/excel.worksheetcollection), [**Worksheet**](https://docs.microsoft.com/javascript/api/excel/excel.worksheet)</span><span class="sxs-lookup"><span data-stu-id="a5a55-120">**WorksheetCollection**, [Worksheet](https://docs.microsoft.com/javascript/api/excel/excel.worksheetcollection)</span></span> |
| `onChanged` | <span data-ttu-id="a5a55-121">Evento que ocorre quando os dados de células são alterados.</span><span class="sxs-lookup"><span data-stu-id="a5a55-121">Event that occurs when data within cells is changed.</span></span> | <span data-ttu-id="a5a55-122">[**Planilha**](https://docs.microsoft.com/javascript/api/excel/excel.worksheet), [**Tabela**](https://docs.microsoft.com/javascript/api/excel/excel.table), [**TableCollection**](https://docs.microsoft.com/javascript/api/excel/excel.tablecollection)</span><span class="sxs-lookup"><span data-stu-id="a5a55-122">[**Worksheet**](https://docs.microsoft.com/javascript/api/excel/excel.worksheet), [**Table**](https://docs.microsoft.com/javascript/api/excel/excel.table), [**TableCollection**](https://docs.microsoft.com/javascript/api/excel/excel.tablecollection)</span></span> |
| `onDataChanged` | <span data-ttu-id="a5a55-123">Evento que ocorre quando os dados ou a formatação dentro da associação são alterados.</span><span class="sxs-lookup"><span data-stu-id="a5a55-123">Occurs when data or formatting within the binding is changed.</span></span> | [<span data-ttu-id="a5a55-124">**Associação**</span><span class="sxs-lookup"><span data-stu-id="a5a55-124">**Binding**</span></span>](https://docs.microsoft.com/javascript/api/excel/excel.binding) |
| `onSelectionChanged` | <span data-ttu-id="a5a55-125">Evento que ocorre quando uma célula ativa ou um intervalo selecionado são alterados.</span><span class="sxs-lookup"><span data-stu-id="a5a55-125">Event that occurs when the active cell or selected range is changed.</span></span> | <span data-ttu-id="a5a55-126">[**Planilha**](https://docs.microsoft.com/javascript/api/excel/excel.worksheet), [**Tabela**](https://docs.microsoft.com/javascript/api/excel/excel.table), [**Associação**](https://docs.microsoft.com/javascript/api/excel/excel.binding)</span><span class="sxs-lookup"><span data-stu-id="a5a55-126">**Worksheet**, [Table](https://docs.microsoft.com/javascript/api/excel/excel.worksheet), **Binding**</span></span> |
| `onSettingsChanged` | <span data-ttu-id="a5a55-127">Evento que ocorre quando as Configurações no documento são alteradas.</span><span class="sxs-lookup"><span data-stu-id="a5a55-127">Occurs when the Settings in the document are changed.</span></span> | [<span data-ttu-id="a5a55-128">**SettingCollection**</span><span class="sxs-lookup"><span data-stu-id="a5a55-128">**settingCollection**</span></span>](https://docs.microsoft.com/javascript/api/excel/excel.settingcollection) |

### <a name="event-triggers"></a><span data-ttu-id="a5a55-129">Gatilhos de eventos</span><span class="sxs-lookup"><span data-stu-id="a5a55-129">Event triggers</span></span>

<span data-ttu-id="a5a55-130">Os eventos em uma pasta de trabalho do Excel podem ser acionados por:</span><span class="sxs-lookup"><span data-stu-id="a5a55-130">Events within an Excel workbook can be triggered by:</span></span>

- <span data-ttu-id="a5a55-131">Interação do usuário por meio da interface do usuário (UI) do Excel que altera a pasta de trabalho</span><span class="sxs-lookup"><span data-stu-id="a5a55-131">User interaction via the Excel user interface (UI) that changes the workbook</span></span>
- <span data-ttu-id="a5a55-132">Código de suplemento do Office (JavaScript) que altera a pasta de trabalho</span><span class="sxs-lookup"><span data-stu-id="a5a55-132">Office add-in (JavaScript) code that changes the workbook</span></span>
- <span data-ttu-id="a5a55-133">Código de suplemento de VBA (macro) que altera a pasta de trabalho</span><span class="sxs-lookup"><span data-stu-id="a5a55-133">VBA add-in (macro) code that changes the workbook</span></span>

<span data-ttu-id="a5a55-134">Todas as alterações que sejam compatíveis com o comportamento padrão do Excel acionarão eventos correspondentes em uma pasta de trabalho.</span><span class="sxs-lookup"><span data-stu-id="a5a55-134">Any change that complies with default behavior of Excel will trigger the corresponding event(s) in a workbook.</span></span>

### <a name="lifecycle-of-an-event-handler"></a><span data-ttu-id="a5a55-135">Ciclo de vida de um manipulador de eventos</span><span class="sxs-lookup"><span data-stu-id="a5a55-135">Lifecycle of an event handler</span></span>

<span data-ttu-id="a5a55-p102">Um manipulador de eventos é criado quando um suplemento registra o manipulador de eventos. Ele é destruído quando o suplemento cancela o registro do manipulador de eventos ou quando os suplementos são atualizados, recarregados ou fechados. Manipuladores de eventos não persistem como parte do arquivo do Excel, ou em sessões com Excel Online.</span><span class="sxs-lookup"><span data-stu-id="a5a55-p102">An event handler is created when an add-in registers the event handler and is destroyed when the add-in unregisters the event handler or when the add-in is closed. Event handlers do not persist as part of the Excel file.</span></span>

> [!CAUTION]
> <span data-ttu-id="a5a55-139">Quando um objeto ao qual os eventos são registrados for excluído (por exemplo, uma tabela com um evento `onChanged` registrado), o manipulador de eventos não dispara mais, porém permanece na memória até o suplemento ou sessão de Excel for atualizado(a) ou fechado(a).</span><span class="sxs-lookup"><span data-stu-id="a5a55-139">When an object to which events are registered is deleted (e.g., a table with an `onChanged` event registered), the event handler no longer triggers but remains in memory until the add-in or Excel session refreshes or closes.</span></span>

### <a name="events-and-coauthoring"></a><span data-ttu-id="a5a55-140">Eventos e coautoria</span><span class="sxs-lookup"><span data-stu-id="a5a55-140">Events and coauthoring</span></span>

<span data-ttu-id="a5a55-p103">Com a [coautoria](co-authoring-in-excel-add-ins.md), várias pessoas podem trabalhar em conjunto e editar a mesma pasta de trabalho do Excel simultaneamente. Em eventos que podem ser acionados por um coautor, como `onChanged`, o objeto **Event** correspondente conterá a propriedade **source** que indica se o evento foi acionado localmente pelo usuário atual (`event.source = Local`) ou pelo coautor remoto (`event.source = Remote`).</span><span class="sxs-lookup"><span data-stu-id="a5a55-p103">With [coauthoring](co-authoring-in-excel-add-ins.md), multiple people can work together and edit the same Excel workbook simultaneously. For events that can be triggered by a coauthor, such as `onChanged`, the corresponding **Event** object will contain a **source** property that indicates whether the event was triggered locally by the current user (`event.source = Local`) or was triggered by the remote coauthor (`event.source = Remote`).</span></span>

## <a name="register-an-event-handler"></a><span data-ttu-id="a5a55-143">Registrar um manipulador de eventos.</span><span class="sxs-lookup"><span data-stu-id="a5a55-143">Register an event handler</span></span>

<span data-ttu-id="a5a55-p104">O exemplo de código a seguir registra um manipulador de eventos para o `onChanged` evento na planilha chamada **Amostra**. O código especifica que, quando dados são alterados nessa planilha, a função `handleDataChange` deverá ser executada.</span><span class="sxs-lookup"><span data-stu-id="a5a55-p104">The following code sample registers an event handler for the `onChanged` event in the worksheet named **Sample**. The code specifies that when data changes in that worksheet, the `handleDataChange` function should run.</span></span>

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

## <a name="handle-an-event"></a><span data-ttu-id="a5a55-146">Manipular um evento</span><span class="sxs-lookup"><span data-stu-id="a5a55-146">Handle an event</span></span>

<span data-ttu-id="a5a55-p105">Conforme mostrado no exemplo anterior, quando você registra um manipulador de eventos, você indica a função que deverá ser executada quando ocorre o evento específico. Você pode projetar aquela função para realizar quaisquer ações que seu cenário exigir. O exemplo de código a seguir mostra uma função de manipulador de eventos que simplesmente escreve informações sobre o evento no console.</span><span class="sxs-lookup"><span data-stu-id="a5a55-p105">As shown in the previous example, when you register an event handler, you indicate the function that should run when the specified event occurs. You can design that function to perform whatever actions your scenario requires. The following code sample shows an event handler function that simply writes information about the event to the console.</span></span> 

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

## <a name="remove-an-event-handler"></a><span data-ttu-id="a5a55-150">Remover um manipulador de eventos</span><span class="sxs-lookup"><span data-stu-id="a5a55-150">Remove an event handler</span></span>

<span data-ttu-id="a5a55-p106">O exemplo de código a seguir registra um manipulador de eventos para o `onSelectionChanged` evento na planilha denominada **Amostra** e define a função `handleSelectionChange` que será executada quando o evento ocorre. Ele também define a função `remove()` que poderá subsequentemente ser chamada para remover o manipulador de eventos.</span><span class="sxs-lookup"><span data-stu-id="a5a55-p106">The following code sample registers an event handler for the `onSelectionChanged` event in the worksheet named **Sample** and defines the `handleSelectionChange` function that will run when the event occurs. It also defines the `remove()` function that can subsequently be called to remove that event handler.</span></span>

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

## <a name="enable-and-disable-events"></a><span data-ttu-id="a5a55-153">Ativar e desativar eventos</span><span class="sxs-lookup"><span data-stu-id="a5a55-153">Enable and disable agents</span></span>

<span data-ttu-id="a5a55-p107">O desempenho de um suplemento pode ser aprimorado por meio da desabilitação de eventos. Por exemplo, seu aplicativo pode nunca precisar receber eventos ou ele poderia ignorar eventos enquanto executa edições de lote de várias entidades.</span><span class="sxs-lookup"><span data-stu-id="a5a55-p107">The performance of an add-in may be improved by disabling events. For example, your app might never need to receive events, or it could ignore events while performing batch-edits of multiple entities.</span></span> 

<span data-ttu-id="a5a55-p108">Eventos são habilitados e desabilitados no nível de [tempo de execução](https://docs.microsoft.com/javascript/api/excel/excel.runtime) . A propriedade `enableEvents` determina se os eventos serão acionados e seus manipuladores serão ativados.</span><span class="sxs-lookup"><span data-stu-id="a5a55-p108">Events are enabled and disabled at the [runtime](https://docs.microsoft.com/javascript/api/excel/excel.runtime) level. The `enableEvents` property determines if events are fired and their handlers are activated.</span></span> 

<span data-ttu-id="a5a55-158">O exemplo de código a seguir mostra como ativar e desativar eventos.</span><span class="sxs-lookup"><span data-stu-id="a5a55-158">The following code sample shows how to toggle events on and off.</span></span>

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

## <a name="see-also"></a><span data-ttu-id="a5a55-159">Confira também</span><span class="sxs-lookup"><span data-stu-id="a5a55-159">See also</span></span>

- [<span data-ttu-id="a5a55-160">Conceitos fundamentais de programação com a API JavaScript do Excel</span><span class="sxs-lookup"><span data-stu-id="a5a55-160">Fundamental programming concepts with the Excel JavaScript API</span></span>](excel-add-ins-core-concepts.md)