---
title: Trabalhar com eventos usando a API JavaScript do Excel
description: Uma lista de eventos para objetos JavaScript do Excel. Isso inclui informações sobre como usar manipuladores de eventos e os padrões associados.
ms.date: 02/11/2020
localization_priority: Normal
ms.openlocfilehash: f1a1faf9acc370e7183a078aeeba34019e54900f
ms.sourcegitcommit: 6c7c98f085dd20f827e0c388e672993412944851
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 03/06/2020
ms.locfileid: "42554784"
---
# <a name="work-with-events-using-the-excel-javascript-api"></a>Trabalhar com eventos usando a API JavaScript do Excel

Este artigo descreve conceitos importantes relacionados ao trabalho com eventos no Excel e fornece exemplos de código que mostram como registrar manipuladores de eventos, lidar com eventos e remover manipuladores de eventos usando as APIs JavaScript do Excel.

## <a name="events-in-excel"></a>Eventos no Excel

Sempre que ocorrerem certos tipos de alterações em uma pasta de trabalho do Excel, uma notificação do evento será ativada. Ao usar as APIs JavaScript do Excel, você pode registrar manipuladores de eventos que permitem que o suplemento execute automaticamente uma função designada quando ocorre um evento específico. Os eventos a seguir têm suporte no momento:

| Evento | Descrição | Objetos com suporte |
|:---------------|:-------------|:-----------|
| `onActivated` | Ocorre quando um objeto está ativado. | [**Gráfico**](/javascript/api/excel/excel.chart), [**ChartCollection**](/javascript/api/excel/excel.chartcollection), [**Shape**](/javascript/api/excel/excel.shape), [**Planilha**](/javascript/api/excel/excel.worksheet), [**WorksheetCollection**](/javascript/api/excel/excel.worksheetcollection) |
| `onAdded` | Ocorre quando um objeto é adicionado à coleção. | [**ChartCollection**](/javascript/api/excel/excel.chartcollection), [**TableCollection**](/javascript/api/excel/excel.tablecollection), [**WorksheetCollection**](/javascript/api/excel/excel.worksheetcollection) |
| `onAutoSaveSettingChanged` | Ocorre quando a `autoSave` configuração é alterada na pasta de trabalho. | [**Workbook**](/javascript/api/excel/excel.workbook) |
| `onCalculated` | Ocorre quando uma planilha terminou um cálculo (ou todas as planilhas do conjunto terminaram). | [**Planilha**](/javascript/api/excel/excel.worksheet), [**WorksheetCollection**](/javascript/api/excel/excel.worksheetcollection) |
| `onChanged` | Ocorre quando os dados das células são alterados. | [**Tabela**](/javascript/api/excel/excel.table), [**TableCollection**](/javascript/api/excel/excel.tablecollection), [**Planilha**](/javascript/api/excel/excel.worksheet), [**WorksheetCollection**](/javascript/api/excel/excel.worksheetcollection) |
| `onColumnSorted` | Ocorre quando uma ou mais colunas são classificadas. Isso acontece como resultado de uma operação de classificação da esquerda para a direita. | [**Planilha**](/javascript/api/excel/excel.worksheet), [**WorksheetCollection**](/javascript/api/excel/excel.worksheetcollection) |
| `onDataChanged` | Ocorre quando os dados ou a formatação dentro da associação são alterados. | [**Associação**](/javascript/api/excel/excel.binding) |
| `onDeactivated` | Ocorre quando um objeto é desativado. | [**Gráfico**](/javascript/api/excel/excel.chart), [**ChartCollection**](/javascript/api/excel/excel.chartcollection), [**Shape**](/javascript/api/excel/excel.shape), [**Planilha**](/javascript/api/excel/excel.worksheet), [**WorksheetCollection**](/javascript/api/excel/excel.worksheetcollection) |
| `onDeleted` | Ocorre quando um objeto é excluído da coleção. | [**ChartCollection**](/javascript/api/excel/excel.chartcollection), [**TableCollection**](/javascript/api/excel/excel.tablecollection), [**WorksheetCollection**](/javascript/api/excel/excel.worksheetcollection) |
| `onFormatChanged` | Ocorre quando o formato é alterado em uma planilha. | [**Planilha**](/javascript/api/excel/excel.worksheet), [**WorksheetCollection**](/javascript/api/excel/excel.worksheetcollection) |
| `onRowSorted` | Ocorre quando uma ou mais linhas são classificadas. Isso ocorre como resultado de uma operação de classificação de cima para baixo. | [**WorksheetCollection**](/javascript/api/excel/excel.worksheet), [**Planilha**](/javascript/api/excel/excel.worksheetcollection) |
| `onSelectionChanged` | Ocorre quando uma célula ativa ou um intervalo selecionado são alterados. | [**Associação**](/javascript/api/excel/excel.binding), [**Tabela**](/javascript/api/excel/excel.table),  [**Worksheet**](/javascript/api/excel/excel.worksheet), [**WorksheetCollection**](/javascript/api/excel/excel.worksheetcollection) |
| `onSettingsChanged` | Ocorre quando as Configurações no documento são alteradas. | [**SettingCollection**](/javascript/api/excel/excel.settingcollection) |
| `onSingleClicked` | Acontece quando a operação é clicada/pressionada com o botão esquerdo do mouse ocorre na planilha. | [**Planilha**](/javascript/api/excel/excel.worksheet), [**WorksheetCollection**](/javascript/api/excel/excel.worksheetcollection) |

> [!WARNING]
> O `onSelectionChanged` atualmente é instável. Existe uma solução alternativa para o uso confiável de `onSelectionChanged`. Adicione o seguinte código à seção `<head>` da sua home page HTML:
>
> ```HTML
> <script> MutationObserver=null; </script>
> ```
>
> Uma discussão completa sobre o assunto pode ser encontrada no [repositório office-js GitHub](https://github.com/OfficeDev/office-js/issues/533).

### <a name="events-in-preview"></a>Eventos no modo de visualização

> [!NOTE]
> Os seguintes eventos estão disponíveis atualmente apenas na visualização pública. [!INCLUDE [Information about using preview APIs](../includes/using-excel-preview-apis.md)]

| Evento | Descrição | Objetos com suporte |
|:---------------|:-------------|:-----------|
| `onFiltered` | Ocorre quando um filtro é aplicado a um objeto. | [**Tabela**](/javascript/api/excel/excel.table), [**TableCollection**](/javascript/api/excel/excel.tablecollection), [**Planilha**](/javascript/api/excel/excel.worksheet), [**WorksheetCollection**](/javascript/api/excel/excel.worksheetcollection) |
| `onRowHiddenChanged` | Ocorre quando o estado de linha oculta é alterado em uma planilha específica. | [**Planilha**](/javascript/api/excel/excel.worksheet), [**WorksheetCollection**](/javascript/api/excel/excel.worksheetcollection) |

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

O exemplo de código a seguir registra um manipulador de eventos para o evento `onSelectionChanged` na planilha **Sample** e define a função `handleSelectionChange` a executar quando o evento ocorrer. Também define a função `remove()` que pode ser chamada posteriormente para remover aquele manipulador de eventos. Observe que o `RequestContext` manipulador de eventos usado para criar o é necessário para removê-lo. 

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

- [Conceitos fundamentais de programação com a API JavaScript do Excel](excel-add-ins-core-concepts.md)
