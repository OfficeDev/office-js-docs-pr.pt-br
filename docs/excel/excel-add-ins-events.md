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
# <a name="work-with-events-using-the-excel-javascript-api"></a>Trabalhar com eventos usando a API JavaScript do Excel 

Este artigo descreve conceitos importantes relacionados ao trabalho com eventos no Excel e fornece exemplos de código que mostram como registrar manipuladores de eventos, lidar com eventos e remover manipuladores de eventos usando as APIs JavaScript do Excel. 

## <a name="events-in-excel"></a>Eventos no Excel

Sempre que ocorrerem certos tipos de alterações em uma pasta de trabalho do Excel, uma notificação do evento será ativada. Ao usar as APIs JavaScript do Excel, você pode registrar manipuladores de eventos que permitem que o suplemento execute automaticamente uma função designada quando ocorre um evento específico. Os eventos a seguir têm suporte no momento:

| Evento | Descrição | Objetos com suporte |
|:---------------|:-------------|:-----------|
| `onAdded` | Evento que ocorre quando um objeto é adicionado. | [**WorksheetCollection**](https://dev.office.com/reference/add-ins/excel/worksheetcollection) |
| `onDeleted` | Evento que ocorre quando um objeto é excluído. | [**WorksheetCollection**](https://dev.office.com/reference/add-ins/excel/worksheetcollection) |
| `onActivated` | Evento que ocorre quando um objeto é ativado. | [**WorksheetCollection**](https://dev.office.com/reference/add-ins/excel/worksheetcollection), [**Planilha**](https://dev.office.com/reference/add-ins/excel/worksheet) |
| `onDeactivated` | Evento que ocorre quando um objeto é desativado. | [**WorksheetCollection**](https://dev.office.com/reference/add-ins/excel/worksheetcollection), [**Planilha**](https://dev.office.com/reference/add-ins/excel/worksheet) |
| `onChanged` | Evento que ocorre quando os dados de células são alterados. | [**Planilha**](https://dev.office.com/reference/add-ins/excel/worksheet), [**Tabela**](https://dev.office.com/reference/add-ins/excel/table), [**TableCollection**](https://dev.office.com/reference/add-ins/excel/tablecollection) |
| `onDataChanged` | Evento que ocorre quando os dados ou a formatação na associação são alterados. | [**Associação**](https://dev.office.com/reference/add-ins/excel/binding) |
| `onSelectionChanged` | Evento que ocorre quando uma célula ativa ou um intervalo selecionado são alterados. | [**Planilha**](https://dev.office.com/reference/add-ins/excel/worksheet), [**Tabela**](https://dev.office.com/reference/add-ins/excel/table), [**Associação**](https://dev.office.com/reference/add-ins/excel/binding) |
| `onSettingsChanged` | Evento que ocorre quando as Configurações no documento são alteradas. | [**SettingCollection**](https://dev.office.com/reference/add-ins/excel/settingcollection) |

## <a name="preview-beta-events-in-excel"></a>Visualizar eventos (beta) no Excel

> [!NOTE]
> Esses eventos estão atualmente disponíveis apenas na versão prévia pública (beta). Para usar esses recursos, você deve usar a biblioteca beta do CDN do Office.js: https://appsforoffice.microsoft.com/lib/beta/hosted/office.js.

| Evento | Descrição | Objetos com suporte |
|:---------------|:-------------|:-----------|
| `onAdded` | Evento que ocorre quando um gráfico é adicionado. | [**ChartCollection**](https://github.com/OfficeDev/office-js-docs/blob/ExcelJs_OpenSpec/reference/new-events.md) |
| `onDeleted` | Evento que ocorre quando um gráfico é excluído. | [**ChartCollection**](https://github.com/OfficeDev/office-js-docs/blob/ExcelJs_OpenSpec/reference/new-events.md) |
| `onActivated` | Evento que ocorre quando um gráfico é ativado. | [**Gráfico**](https://github.com/OfficeDev/office-js-docs/blob/ExcelJs_OpenSpec/reference/new-events.md), [**ChartCollection**](https://github.com/OfficeDev/office-js-docs/blob/ExcelJs_OpenSpec/reference/new-events.md) |
| `onDeactivated` | Evento que ocorre quando um gráfico é desativado. | [**Gráfico**](https://github.com/OfficeDev/office-js-docs/blob/ExcelJs_OpenSpec/reference/new-events.md), [**ChartCollection**](https://github.com/OfficeDev/office-js-docs/blob/ExcelJs_OpenSpec/reference/new-events.md) |
| `onCalculated` | Evento que ocorre quando uma planilha termina o cálculo (ou todas as planilhas da coleção foram concluídas). | [**WorksheetCollection**](https://github.com/OfficeDev/office-js-docs/blob/ExcelJs_OpenSpec/reference/new-events.md), [**Planilha**](https://github.com/OfficeDev/office-js-docs/blob/ExcelJs_OpenSpec/reference/new-events.md) |

### <a name="event-triggers"></a>Gatilhos de eventos

Os eventos em uma pasta de trabalho do Excel podem ser acionados por:

- Interação do usuário por meio da interface do usuário (UI) do Excel que altere a pasta de trabalho
- Código de suplemento do Office (em JavaScript) que altere a pasta de trabalho
- Código de suplemento de VBA (macro) que altere a pasta de trabalho

Todas as alterações que sejam compatíveis com o comportamento padrão do Excel acionarão eventos correspondentes em uma pasta de trabalho.

### <a name="lifecycle-of-an-event-handler"></a>Ciclo de vida de um manipulador de eventos

Um manipulador de eventos é criado quando um suplemento o registra e é destruído quando o suplemento cancela seu registro ou quando o suplemento for fechado. Os manipuladores de eventos não persistem como parte do arquivo de Excel.

### <a name="events-and-coauthoring"></a>Eventos e coautoria

Com a [coautoria](co-authoring-in-excel-add-ins.md), várias pessoas podem trabalhar em conjunto e editar a mesma pasta de trabalho do Excel simultaneamente. Em eventos que podem ser disparados por um coautor, como `onChanged`, o objeto de **evento** respectivo conterá a propriedade **fonte** que indica se o evento foi acionado localmente pelo usuário atual (`event.source = Local`) ou pelo coautor remoto (`event.source = Remote`).

## <a name="register-an-event-handler"></a>Registrar um manipulador de eventos.

O exemplo de código a seguir registra um manipulador de eventos para o evento `onChanged` na planilha **Sample**. O código especifica que, quando os dados forem alterados na planilha, a função `handleDataChange` deve ser executada.

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

Como mostrado no exemplo anterior, quando você registrar um manipulador de eventos, indica a função a ser executada quando o evento especificado ocorre. Você pode criar essa função para executar as ações que seu cenário exige. O exemplo de código a seguir mostra uma função de manipulador de eventos que simplesmente grava informações sobre o evento no console. 

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

O exemplo de código a seguir registra um manipulador de eventos para o evento `onSelectionChanged` na planilha **Sample** e define a função `handleSelectionChange` a executar quando o evento ocorrer. Também define a função `remove()` que pode ser chamada posteriormente para remover aquele manipulador de eventos.

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

## <a name="enable-and-disable-events"></a>Ativar e desativar eventos

> [!NOTE]
> Este recurso está atualmente disponível somente na versão prévia pública (beta). Para usá-lo, você deve fazer referência a biblioteca beta da CDN do Office. js: https://appsforoffice.microsoft.com/lib/beta/hosted/office.js.

Eventos são ativados e desativados em [tempo de execução](https://docs.microsoft.com/en-us/javascript/api/excel/excel.runtime?view=office-js). A propriedade `enableEvents` determina se os eventos são disparados e seus manipuladores serão ativados. Desativar eventos é útil quando o desempenho é crítico ou durante a edição de várias entidades e se deseja evitar o acionamento de eventos até terminar.

O exemplo de código a seguir mostra como ativar e desativar eventos.

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

## <a name="see-also"></a>Veja também

- [Principais conceitos da API JavaScript do Excel](excel-add-ins-core-concepts.md)
- [Especificação para abrir API JavaScript do Excel](https://github.com/OfficeDev/office-js-docs/tree/ExcelJs_OpenSpec)