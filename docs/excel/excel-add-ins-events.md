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
# <a name="work-with-events-using-the-excel-javascript-api"></a>Trabalhar com eventos usando a API JavaScript do Excel 

Este artigo descreve conceitos importantes relacionados ao trabalho com eventos no Excel e fornece exemplos de c?digo que mostram como registrar manipuladores de eventos, lidar com eventos e remover manipuladores de eventos usando as APIs JavaScript do Excel. 

## <a name="events-in-excel"></a>Eventos no Excel

Sempre que ocorrerem certos tipos de altera??es em uma pasta de trabalho do Excel, uma notifica??o do evento ser? ativada. Ao usar as APIs JavaScript do Excel, voc? pode registrar manipuladores de eventos que permitem que o suplemento execute automaticamente uma fun??o designada quando ocorre um evento espec?fico. Os eventos a seguir t?m suporte no momento:

| Evento | Descri??o | Objetos com suporte |
|:---------------|:-------------|:-----------|
| `onAdded` | Evento que ocorre quando um objeto ? adicionado. | [**WorksheetCollection**](https://dev.office.com/reference/add-ins/excel/worksheetcollection) |
| `onDeleted` | Evento que ocorre quando um objeto ? exclu?do. | [**WorksheetCollection**](https://dev.office.com/reference/add-ins/excel/worksheetcollection) |
| `onActivated` | Evento que ocorre quando um objeto ? ativado. | [**WorksheetCollection**](https://dev.office.com/reference/add-ins/excel/worksheetcollection), [**Planilha**](https://dev.office.com/reference/add-ins/excel/worksheet) |
| `onDeactivated` | Evento que ocorre quando um objeto ? desativado. | [**WorksheetCollection**](https://dev.office.com/reference/add-ins/excel/worksheetcollection), [**Planilha**](https://dev.office.com/reference/add-ins/excel/worksheet) |
| `onChanged` | Evento que ocorre quando os dados das c?lulas s?o alterados. | [**Planilha**](https://dev.office.com/reference/add-ins/excel/worksheet), [**Tabela**](https://dev.office.com/reference/add-ins/excel/table), [**TableCollection**](https://dev.office.com/reference/add-ins/excel/tablecollection) |
| `onDataChanged` | Evento que ocorre quando os dados ou a formata??o dentro da associa??o s?o alterados. | [**Associa??o**](https://dev.office.com/reference/add-ins/excel/binding) |
| `onSelectionChanged` | Evento que ocorre quando uma c?lula ativa ou um intervalo selecionado s?o alterados. | [**Planilha**](https://dev.office.com/reference/add-ins/excel/worksheet), [**Tabela**](https://dev.office.com/reference/add-ins/excel/table), [**Associa??o**](https://dev.office.com/reference/add-ins/excel/binding) |

### <a name="event-triggers"></a>Gatilhos de eventos

Os eventos em uma pasta de trabalho do Excel podem ser acionados por:

- Intera??o do usu?rio por meio da interface do usu?rio (UI) do Excel que altere a pasta de trabalho
- C?digo de suplemento do Office (em JavaScript) que altere a pasta de trabalho
- C?digo de suplemento de VBA (macro) que altere a pasta de trabalho

Todas as altera??es que sejam compat?veis com o comportamento padr?o do Excel acionar?o eventos correspondentes em uma pasta de trabalho.

### <a name="lifecycle-of-an-event-handler"></a>Ciclo de vida de um manipulador de eventos

Um manipulador de eventos ? criado quando um suplemento o registra e ? destru?do quando o suplemento cancela seu registro ou quando o suplemento for fechado. Os manipuladores de eventos n?o persistem como parte do arquivo de Excel.

### <a name="events-and-coauthoring"></a>Eventos e coautoria

Com a [coautoria](co-authoring-in-excel-add-ins.md), v?rias pessoas podem trabalhar em conjunto e editar a mesma pasta de trabalho do Excel simultaneamente. Em eventos que podem ser disparados por um coautor, como `onChanged`, o objeto de **evento** respectivo conter? a propriedade **fonte** que indica se o evento foi acionado localmente pelo usu?rio atual (`event.source = Local`) ou pelo coautor remoto (`event.source = Remote`).

## <a name="register-an-event-handler"></a>Registrar um manipulador de eventos.

O exemplo de c?digo a seguir registra um manipulador de eventos para o evento `onChanged` na planilha **Sample**. O c?digo especifica que, quando os dados forem alterados na planilha, a fun??o `handleDataChange` deve ser executada.

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

## <a name="handle-an-event"></a>Manipular um evento

Como mostrado no exemplo anterior, quando voc? registrar um manipulador de eventos, indica a fun??o a ser executada quando o evento especificado ocorre. Voc? pode criar essa fun??o para executar as a??es que seu cen?rio exige. O exemplo de c?digo a seguir mostra uma fun??o de manipulador de eventos que simplesmente grava informa??es sobre o evento no console. 

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

## <a name="remove-an-event-handler"></a>Remover um manipulador de eventos

O exemplo de c?digo a seguir registra um manipulador de eventos para o evento `onSelectionChanged` na planilha **Sample** e define a fun??o `handleSelectionChange` a executar quando o evento ocorrer. Tamb?m define a fun??o `remove()` que pode ser chamada posteriormente para remover aquele manipulador de eventos.

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

## <a name="see-also"></a>Confira tamb?m

- [Principais conceitos da API JavaScript do Excel](excel-add-ins-core-concepts.md)
- [Especifica??o para abrir API JavaScript do Excel](https://github.com/OfficeDev/office-js-docs/tree/ExcelJs_OpenSpec)