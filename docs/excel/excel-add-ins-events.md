---
title: Trabalhar com eventos usando a API JavaScript do Excel
description: Uma lista de eventos para Excel JavaScript. Isso inclui informações sobre como usar manipuladores de eventos e os padrões associados.
ms.date: 07/02/2021
localization_priority: Normal
ms.openlocfilehash: e908a9253649a47838e762f03b930838115847c5927333f3af82bd00bdc90829
ms.sourcegitcommit: 4f2c76b48d15e7d03c5c5f1f809493758fcd88ec
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 08/07/2021
ms.locfileid: "57085505"
---
# <a name="work-with-events-using-the-excel-javascript-api"></a>Trabalhar com eventos usando a API JavaScript do Excel

Este artigo descreve conceitos importantes relacionados ao trabalho com eventos no Excel e fornece exemplos de código que mostram como registrar manipuladores de eventos, lidar com eventos e remover manipuladores de eventos usando as APIs JavaScript do Excel.

## <a name="events-in-excel"></a>Eventos no Excel

Sempre que ocorrerem certos tipos de alterações em uma pasta de trabalho do Excel, uma notificação do evento será ativada. Ao usar as APIs JavaScript do Excel, você pode registrar manipuladores de eventos que permitem que o suplemento execute automaticamente uma função designada quando ocorre um evento específico. Os eventos a seguir têm suporte no momento:

| Evento | Descrição | Objetos com suporte |
|:---------------|:-------------|:-----------|
| `onActivated` | Ocorre quando um objeto está ativado. | [**Gráfico**](/javascript/api/excel/excel.chart#onActivated), [**ChartCollection**](/javascript/api/excel/excel.chartcollection#onActivated), [**Shape**](/javascript/api/excel/excel.shape#onActivated), [**Planilha**](/javascript/api/excel/excel.worksheet#onActivated), [**WorksheetCollection**](/javascript/api/excel/excel.worksheetcollection#onActivated) |
| `onActivated` | Ocorre quando uma workbook é ativada. | [**Workbook**](/javascript/api/excel/excel.workbook#onActivated) |
| `onAdded` | Ocorre quando um objeto é adicionado à coleção. | [**ChartCollection,**](/javascript/api/excel/excel.chartcollection#onAdded) [**CommentCollection,**](/javascript/api/excel/excel.commentcollection#onAdded) [**TableCollection,**](/javascript/api/excel/excel.tablecollection#onAdded) [**WorksheetCollection**](/javascript/api/excel/excel.worksheetcollection#onAdded) |
| `onAutoSaveSettingChanged` | Ocorre quando a `autoSave` configuração é alterada na pasta de trabalho. | [**Workbook**](/javascript/api/excel/excel.workbook#onAutoSaveSettingChanged) |
| `onCalculated` | Ocorre quando uma planilha terminou um cálculo (ou todas as planilhas do conjunto terminaram). | [**WorksheetCollection**](/javascript/api/excel/excel.worksheet#onCalculated), [**Planilha**](/javascript/api/excel/excel.worksheetcollection#onCalculated) |
| `onChanged` | Ocorre quando os dados de células individuais ou comentários foram alterados. | [**CommentCollection**](/javascript/api/excel/excel.commentcollection#onChanged), [**Table**](/javascript/api/excel/excel.table#onChanged), [**TableCollection**](/javascript/api/excel/excel.tablecollection#onChanged), [**Worksheet**](/javascript/api/excel/excel.worksheet#onChanged), [**WorksheetCollection**](/javascript/api/excel/excel.worksheetcollection#onChanged) |
| `onColumnSorted` | Ocorre quando uma ou mais colunas são classificadas. Isso acontece como resultado de uma operação de classificação da esquerda para a direita. | [**Planilha**](/javascript/api/excel/excel.worksheet#onColumnSorted), [**WorksheetCollection**](/javascript/api/excel/excel.worksheetcollection#onColumnSorted) |
| `onDataChanged` | Ocorre quando os dados ou a formatação dentro da associação são alterados. | [**Associação**](/javascript/api/excel/excel.binding#onDataChanged) |
| `onDeactivated` | Ocorre quando um objeto é desativado. | [**Gráfico**](/javascript/api/excel/excel.chart#onDeactivated), [**ChartCollection**](/javascript/api/excel/excel.chartcollection#onDeactivated), [**Shape**](/javascript/api/excel/excel.shape#onDeactivated), [**Planilha**](/javascript/api/excel/excel.worksheet#onDeactivated), [**WorksheetCollection**](/javascript/api/excel/excel.worksheetcollection#onDeactivated) |
| `onDeleted` | Ocorre quando um objeto é excluído da coleção. | [**ChartCollection,**](/javascript/api/excel/excel.chartcollection#onDeleted) [**CommentCollection,**](/javascript/api/excel/excel.commentcollection#onDeleted) [**TableCollection,**](/javascript/api/excel/excel.tablecollection#onDeleted) [**WorksheetCollection**](/javascript/api/excel/excel.worksheetcollection#onDeleted) |
| `onFormatChanged` | Ocorre quando o formato é alterado em uma planilha. | [**Planilha**](/javascript/api/excel/excel.worksheet#onFormatChanged), [**WorksheetCollection**](/javascript/api/excel/excel.worksheetcollection#onFormatChanged) |
| `onFormulaChanged` | Ocorre quando uma fórmula é alterada. | [**Planilha**](/javascript/api/excel/excel.worksheet#onFormulaChanged), [**WorksheetCollection**](/javascript/api/excel/excel.worksheetcollection#onFormulaChanged) |
| `onRowSorted` | Ocorre quando uma ou mais linhas são classificadas. Isso ocorre como resultado de uma operação de classificação de cima para baixo. | [**Planilha**](/javascript/api/excel/excel.worksheet#onRowSorted), [**WorksheetCollection**](/javascript/api/excel/excel.worksheetcollection#onRowSorted) |
| `onSelectionChanged` | Ocorre quando uma célula ativa ou um intervalo selecionado são alterados. | [**Binding**](/javascript/api/excel/excel.binding#onSelectionChanged), [**Table**](/javascript/api/excel/excel.table#onSelectionChanged), [**Workbook**](/javascript/api/excel/excel.workbook#onSelectionChanged), [**Worksheet**](/javascript/api/excel/excel.worksheet#onSelectionChanged), [**WorksheetCollection**](/javascript/api/excel/excel.worksheetcollection#onSelectionChanged) |
| `onRowHiddenChanged` | Ocorre quando o estado de linha oculta é alterado em uma planilha específica. | [**Planilha**](/javascript/api/excel/excel.worksheet#onRowHiddenChanged), [**WorksheetCollection**](/javascript/api/excel/excel.worksheetcollection#onRowHiddenChanged) |
| `onSettingsChanged` | Ocorre quando as Configurações no documento são alteradas. | [**SettingCollection**](/javascript/api/excel/excel.settingcollection#onSettingsChanged) |
| `onSingleClicked` | Acontece quando a operação é clicada/pressionada com o botão esquerdo do mouse ocorre na planilha. | [**Planilha**](/javascript/api/excel/excel.worksheet#onSingleClicked), [**WorksheetCollection**](/javascript/api/excel/excel.worksheetcollection#onSingleClicked) |

### <a name="events-in-preview"></a>Eventos no modo de visualização

> [!NOTE]
> Os seguintes eventos estão disponíveis atualmente apenas na visualização pública. [!INCLUDE [Information about using preview APIs](../includes/using-excel-preview-apis.md)]

| Evento | Descrição | Objetos com suporte |
|:---------------|:-------------|:-----------|
| `onFiltered` | Ocorre quando um filtro é aplicado a um objeto. | [**Tabela**](/javascript/api/excel/excel.table#onFiltered), [**TableCollection**](/javascript/api/excel/excel.tablecollection#onFiltered), [**Planilha**](/javascript/api/excel/excel.worksheet#onFiltered), [**WorksheetCollection**](/javascript/api/excel/excel.worksheetcollection#onFiltered) |

### <a name="event-triggers"></a>Gatilhos de eventos

Os eventos em uma pasta de trabalho do Excel podem ser acionados por:

- Interação do usuário por meio da interface (IU) do Excel que altera a pasta de trabalho
- Código de suplemento do Office (JavaScript) que altera a pasta de trabalho
- Código de suplemento VBA (macro) que altera a pasta de trabalho

Todas as alterações que sejam compatíveis com o comportamento padrão do Excel acionarão eventos correspondentes em uma pasta de trabalho.

### <a name="lifecycle-of-an-event-handler"></a>Ciclo de vida de um manipulador de eventos

Um manipulador de eventos é criado quando um suplemento registra o manipulador de eventos. Ele é destruído quando o suplemento cancela o registro de manipulador de evento ou quando o suplemento é atualizado, recarregado ou fechado. Manipuladores de eventos não são mantidos como parte do arquivo do Excel ou em sessões do Excel Online.

> [!CAUTION]
> Quando um objeto ao qual os eventos são registrados é excluído (por exemplo, uma tabela com um `onChanged` evento registrado), o manipulador de eventos não disparará mais, mas permanecerá na memória até que o suplemento ou sessão do Excel atualize ou feche.

### <a name="events-and-coauthoring"></a>Eventos e coautoria

Com a [coautoria](co-authoring-in-excel-add-ins.md), várias pessoas podem trabalhar em conjunto e editar a mesma pasta de trabalho do Excel simultaneamente. Em eventos que podem ser disparados por um coautor, como `onChanged`, o objeto de **evento** respectivo conterá a propriedade **fonte** que indica se o evento foi acionado localmente pelo usuário atual (`event.source = Local`) ou pelo coautor remoto (`event.source = Remote`).

## <a name="register-an-event-handler"></a>Registrar um manipulador de eventos.

O exemplo de código a seguir registra um manipulador de eventos para o evento `onChanged` na planilha **Sample**. O código especifica que, quando os dados forem alterados na planilha, a função `handleChange` deve ser executada.

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

O exemplo de código a seguir registra um manipulador de eventos para o evento `onSelectionChanged` na planilha **Sample** e define a função `handleSelectionChange` a executar quando o evento ocorrer. Também define a função `remove()` que pode ser chamada posteriormente para remover aquele manipulador de eventos. Observe que o `RequestContext` usado para criar o manipulador de eventos é necessário para removê-lo. 

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

## <a name="enable-and-disable-events"></a>Habilitar e desabilitar eventos

O desempenho de um suplemento pode ser melhorado desabilitando eventos.
Por exemplo, seu aplicativo pode não precisar receber eventos ou ele pode ignorar eventos durante a edições de lotes de várias entidades.

Os eventos são habilitados ou desabilitados no nível [runtime](/javascript/api/excel/excel.runtime).
A `enableEvents` propriedade determina se os eventos são disparados e se seus manipuladores são ativados.

O código a seguir mostra como ativar ou desativar os eventos.

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

- [Modelo de objeto JavaScript do Excel em Suplementos do Office](excel-add-ins-core-concepts.md)
