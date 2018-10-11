---
title: Trabalhar com eventos usando a API JavaScript do Excel
description: ''
ms.date: 09/21/2018
ms.openlocfilehash: b56d25e7e0306b4881115397d4136e63ddc03e5c
ms.sourcegitcommit: 563c53bac52b31277ab935f30af648f17c5ed1e2
ms.translationtype: HT
ms.contentlocale: pt-BR
ms.lasthandoff: 10/10/2018
ms.locfileid: "25459172"
---
# <a name="work-with-events-using-the-excel-javascript-api"></a>Trabalhar com eventos usando a API JavaScript do Excel 

Este artigo descreve conceitos importantes relacionados ao trabalho com eventos no Excel e fornece exemplos de código que mostram como registrar manipuladores de eventos, lidar com eventos e remover manipuladores de eventos usando as APIs JavaScript do Excel. 

## <a name="events-in-excel"></a>Eventos no Excel

Sempre que ocorrerem determinados tipos de alterações em uma pasta de trabalho do Excel, uma notificação de evento é acionada. Usando a API JavaScript do Excel, você pode registrar manipuladores de eventos que permitem o suplemento executar automaticamente uma função designada, quando ocorre um evento específico. Os eventos a seguir são suportados no momento.

| Evento | Descrição | Objetos com suporte |
|:---------------|:-------------|:-----------|
| `onAdded` | Evento que ocorre quando um objeto é adicionado. | [**ChartCollection**](https://docs.microsoft.com/javascript/api/excel/excel.chartcollection), [**WorksheetCollection**](https://docs.microsoft.com/javascript/api/excel/excel.worksheetcollection) |
| `onDeleted` | Evento que ocorre quando um objeto é excluído. | [**ChartCollection**](https://docs.microsoft.com/javascript/api/excel/excel.chartcollection), [**WorksheetCollection**](https://docs.microsoft.com/javascript/api/excel/excel.worksheetcollection) |
| `onActivated` | Evento que ocorre quando um objeto é ativado. | [**Gráfico**](https://docs.microsoft.com/javascript/api/excel/excel.chart), [**ChartCollection**](https://docs.microsoft.com/javascript/api/excel/excel.chartcollection), [**WorksheetCollection**](https://docs.microsoft.com/javascript/api/excel/excel.worksheetcollection), [**Planilha**](https://docs.microsoft.com/javascript/api/excel/excel.worksheet) |
| `onDeactivated` | Evento que ocorre quando um objeto é desativado. | [**Gráfico**](https://docs.microsoft.com/javascript/api/excel/excel.chart), [**ChartCollection**](https://docs.microsoft.com/javascript/api/excel/excel.chartcollection), [**WorksheetCollection**](https://docs.microsoft.com/javascript/api/excel/excel.worksheetcollection), [**Planilha**](https://docs.microsoft.com/javascript/api/excel/excel.worksheet) |
| `onCalculated` | Evento que ocorre quando uma planilha termina o cálculo (ou todas as planilhas da coleção foram concluídas). | [**WorksheetCollection**](https://docs.microsoft.com/javascript/api/excel/excel.worksheetcollection), [**Worksheet**](https://docs.microsoft.com/javascript/api/excel/excel.worksheet) |
| `onChanged` | Evento que ocorre quando os dados de células são alterados. | [**Planilha**](https://docs.microsoft.com/javascript/api/excel/excel.worksheet), [**Tabela**](https://docs.microsoft.com/javascript/api/excel/excel.table), [**TableCollection**](https://docs.microsoft.com/javascript/api/excel/excel.tablecollection) |
| `onDataChanged` | Evento que ocorre quando os dados ou a formatação dentro da associação são alterados. | [**Associação**](https://docs.microsoft.com/javascript/api/excel/excel.binding) |
| `onSelectionChanged` | Evento que ocorre quando uma célula ativa ou um intervalo selecionado são alterados. | [**Planilha**](https://docs.microsoft.com/javascript/api/excel/excel.worksheet), [**Tabela**](https://docs.microsoft.com/javascript/api/excel/excel.table), [**Associação**](https://docs.microsoft.com/javascript/api/excel/excel.binding) |
| `onSettingsChanged` | Evento que ocorre quando as Configurações no documento são alteradas. | [**SettingCollection**](https://docs.microsoft.com/javascript/api/excel/excel.settingcollection) |

### <a name="event-triggers"></a>Gatilhos de eventos

Os eventos em uma pasta de trabalho do Excel podem ser acionados por:

- Interação do usuário por meio da interface do usuário (UI) do Excel que altera a pasta de trabalho
- Código de suplemento do Office (JavaScript) que altera a pasta de trabalho
- Código de suplemento de VBA (macro) que altera a pasta de trabalho

Todas as alterações que sejam compatíveis com o comportamento padrão do Excel acionarão eventos correspondentes em uma pasta de trabalho.

### <a name="lifecycle-of-an-event-handler"></a>Ciclo de vida de um manipulador de eventos

Um manipulador de eventos é criado quando um suplemento o registra e é destruído quando o suplemento cancela seu registro ou quando o suplemento for fechado. Os manipuladores de eventos não persistem como parte do arquivo de Excel.

### <a name="events-and-coauthoring"></a>Eventos e coautoria

Com a [coautoria](co-authoring-in-excel-add-ins.md), várias pessoas podem trabalhar em conjunto e editar a mesma pasta de trabalho do Excel simultaneamente. Em eventos que podem ser acionados por um coautor, como `onChanged`, o objeto **Event** correspondente conterá a propriedade **source** que indica se o evento foi acionado localmente pelo usuário atual (`event.source = Local`) ou pelo coautor remoto (`event.source = Remote`).

## <a name="register-an-event-handler"></a>Registrar um manipulador de eventos.

O exemplo de código a seguir registra um manipulador de eventos para o `onChanged` evento na planilha chamada **Amostra**. O código especifica que, quando dados são alterados nessa planilha, a função `handleDataChange` deverá ser executada.

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

Conforme mostrado no exemplo anterior, quando você registra um manipulador de eventos, você indica a função que deverá ser executada quando ocorre o evento específico. Você pode projetar aquela função para realizar quaisquer ações que seu cenário exigir. O exemplo de código a seguir mostra uma função de manipulador de eventos que simplesmente escreve informações sobre o evento no console. 

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

O exemplo de código a seguir registra um manipulador de eventos para o `onSelectionChanged` evento na planilha denominada **Amostra** e define a função `handleSelectionChange` que será executada quando o evento ocorre. Ele também define a função `remove()` que poderá subsequentemente ser chamada para remover o manipulador de eventos.

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

O desempenho de um suplemento pode ser aprimorado por meio da desabilitação de eventos. Por exemplo, seu aplicativo pode nunca precisar receber eventos ou ele poderia ignorar eventos enquanto executa edições de lote de várias entidades. 

Eventos são habilitados e desabilitados no nível de [tempo de execução](https://docs.microsoft.com/javascript/api/excel/excel.runtime) . A propriedade `enableEvents` determina se os eventos serão acionados e seus manipuladores serão ativados. 

O exemplo de código a seguir mostra como ativar e desativar eventos.

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

## <a name="see-also"></a>Confira também

- [Conceitos de programação fundamentais com a API JavaScript do Excel](excel-add-ins-core-concepts.md)