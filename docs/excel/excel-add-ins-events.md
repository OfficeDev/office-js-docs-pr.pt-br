---
title: Trabalhar com eventos usando a API JavaScript do Excel
description: ''
ms.date: 01/29/2018
---

# <a name="work-with-events-using-the-excel-javascript-api"></a>Trabalhar com eventos usando a API JavaScript do Excel

Este artigo descreve conceitos importantes relacionados ao trabalho com eventos no Excel e fornece exemplos de código que mostram como registrar manipuladores de eventos, lidar com eventos e remover manipuladores de eventos usando as APIs JavaScript do Excel. 

> [!IMPORTANT]
> As APIs descritas neste artigo no momento estão disponíveis somente em visualização pública (beta) e não se destinam à utilização em ambientes de produção. Para executar as amostras de código contidas neste artigo, use uma versão recente do Office e faça referências à biblioteca beta do CDN Office.js: https://appsforoffice.microsoft.com/lib/beta/hosted/office.js.

## <a name="events-in-excel"></a>Eventos no Excel

Sempre que ocorrerem certos tipos de alterações em uma pasta de trabalho do Excel, uma notificação do evento será ativada. Ao usar as APIs JavaScript do Excel, você pode registrar manipuladores de eventos que permitem que o suplemento execute automaticamente uma função designada quando ocorre um evento específico. Os eventos a seguir têm suporte no momento:

| Evento | Descrição | Objetos com suporte |
|:---------------|:-------------|:-----------|
| `onAdded` | Evento que ocorre quando um objeto é adicionado. | **WorksheetCollection** |
| `onDeleted`  | Evento que ocorre quando um objeto é excluído. | **WorksheetCollection** |
| `onActivated` | Evento que ocorre quando um objeto é ativado. | **WorksheetCollection**, **Planilha** |
| `onDeactivated` | Evento que ocorre quando um objeto é desativado. | **WorksheetCollection**, **Planilha** |
| `onDataChanged` | Evento que ocorre quando os dados de células são alterados. | **Planilha**, **tabela**, **TableCollection**, **associação** |
| `onSelectionChanged` | Evento que ocorre quando uma célula ativa ou um intervalo selecionado são alterados. | **Planilha**, **tabela**, **associação** |

### <a name="event-triggers"></a>Gatilhos de eventos

Os eventos em uma pasta de trabalho do Excel podem ser acionados por:

- Interação do usuário por meio da interface do usuário (UI) do Excel que altere a pasta de trabalho
- Código de suplemento do Office (em JavaScript) que altere a pasta de trabalho
- Código de suplemento de VBA (macro) que altere a pasta de trabalho

Todas as alterações que sejam compatíveis com o comportamento padrão do Excel acionarão eventos correspondentes em uma pasta de trabalho.

### <a name="lifecycle-of-an-event-handler"></a>Ciclo de vida de um manipulador de eventos

Um manipulador de eventos é criado quando um suplemento registra o manipulador de eventos e ele é destruído quando o suplemento cancela o registro de manipulador de eventos ou quando o suplemento for fechado. Manipuladores de eventos não persistem como parte do arquivo do Excel.

### <a name="events-and-coauthoring"></a>Eventos e coautoria

Com a [coautoria](co-authoring-in-excel-add-ins.md), várias pessoas podem trabalhar juntas e editar simultaneamente a mesma pasta de trabalho do Excel. Em eventos que podem ser disparados por um coautor, como `onDataChanged`, o objeto de **evento** correspondente conterá a propriedade **fonte** que indica se o evento foi acionado localmente pelo usuário atual (`event.source = Local`) ou se foi acionado pelo coautor remoto (`event.source = Remote`).

## <a name="register-an-event-handler"></a>Registrar um manipulador de eventos.

O exemplo de código a seguir registra um manipulador de eventos para o evento `onDataChanged` na planilha **Sample**. O código especifica que, quando os dados forem alterados na planilha, a função `handleDataChange` deve ser executada.

```js
Excel.run(function (context) {
    var worksheet = context.workbook.worksheets.getItem("Sample");
    worksheet.onDataChanged.add(handleDataChange);

    return context.sync()
        .then(function () {
            console.log("Event handler successfully registered for onDataChanged event in the worksheet.");
        });
}).catch(errorHandlerFunction);
```

## <a name="handle-an-event"></a>Manipular um evento

Como mostrado no exemplo anterior, quando você registrar um manipulador de eventos, indica a função a ser executada quando o evento especificado ocorre. Você pode criar essa função para executar as ações que seu cenário exige. O exemplo de código a seguir mostra uma função de manipulador de eventos que simplesmente grava informações sobre o evento no console. 

```js
function handleDataChange(event)
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

## <a name="see-also"></a>Confira também

- [Principais conceitos da API JavaScript do Excel](excel-add-ins-core-concepts.md)
- [Especificação para abrir API JavaScript do Excel](https://github.com/OfficeDev/office-js-docs/tree/ExcelJs_OpenSpec)
- [Introdução a recursos de eventos do Excel (versão prévia)](https://github.com/OfficeDev/office-js-docs/blob/ExcelJs_OpenSpec/Event_README.md)