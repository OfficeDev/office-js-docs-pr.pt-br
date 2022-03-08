---
title: Trabalhar com eventos usando a API JavaScript do Excel
description: Uma lista de eventos para Excel objetos JavaScript. Isso inclui informações sobre como usar manipuladores de eventos e os padrões associados.
ms.date: 02/16/2022
ms.localizationpriority: medium
ms.openlocfilehash: 8bc1dcad8bccb51dbcedfee741954fabf6967670
ms.sourcegitcommit: 7b6ee73fa70b8e0ff45c68675dd26dd7a7b8c3e9
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 03/08/2022
ms.locfileid: "63340579"
---
# <a name="work-with-events-using-the-excel-javascript-api"></a>Trabalhar com eventos usando a API JavaScript do Excel

Este artigo descreve conceitos importantes relacionados ao trabalho com eventos no Excel e fornece exemplos de código que mostram como registrar manipuladores de eventos, lidar com eventos e remover manipuladores de eventos usando as APIs JavaScript do Excel.

## <a name="events-in-excel"></a>Eventos no Excel

Sempre que ocorrerem certos tipos de alterações em uma pasta de trabalho do Excel, uma notificação do evento será ativada. Ao usar as APIs JavaScript do Excel, você pode registrar manipuladores de eventos que permitem que o suplemento execute automaticamente uma função designada quando ocorre um evento específico. Os eventos a seguir têm suporte no momento:

| Evento | Descrição | Objetos com suporte |
|:---------------|:-------------|:-----------|
| `onActivated` | Ocorre quando um objeto está ativado. | [**Gráfico**](/javascript/api/excel/excel.chart#excel-excel-chart-onactivated-member), [**ChartCollection**](/javascript/api/excel/excel.chartcollection#excel-excel-chartcollection-onactivated-member), [**Shape**](/javascript/api/excel/excel.shape#excel-excel-shape-onactivated-member), [**Planilha**](/javascript/api/excel/excel.worksheet#excel-excel-worksheet-onactivated-member), [**WorksheetCollection**](/javascript/api/excel/excel.worksheetcollection#excel-excel-worksheetcollection-onactivated-member) |
| `onActivated` | Ocorre quando uma workbook é ativada. | [**Workbook**](/javascript/api/excel/excel.workbook#excel-excel-workbook-onactivated-member) |
| `onAdded` | Ocorre quando um objeto é adicionado à coleção. | [**ChartCollection**](/javascript/api/excel/excel.chartcollection#excel-excel-chartcollection-onadded-member), [**CommentCollection**](/javascript/api/excel/excel.commentcollection#excel-excel-commentcollection-onadded-member), [**TableCollection**](/javascript/api/excel/excel.tablecollection#excel-excel-tablecollection-onadded-member), [**WorksheetCollection**](/javascript/api/excel/excel.worksheetcollection#excel-excel-worksheetcollection-onadded-member) |
| `onAutoSaveSettingChanged` | Ocorre quando a `autoSave` configuração é alterada na pasta de trabalho. | [**Workbook**](/javascript/api/excel/excel.workbook#excel-excel-workbook-onautosavesettingchanged-member) |
| `onCalculated` | Ocorre quando uma planilha terminou um cálculo (ou todas as planilhas do conjunto terminaram). | [**WorksheetCollection**](/javascript/api/excel/excel.worksheet#excel-excel-worksheet-oncalculated-member), [**Planilha**](/javascript/api/excel/excel.worksheetcollection#excel-excel-worksheetcollection-oncalculated-member) |
| `onChanged` | Ocorre quando os dados de células individuais ou comentários foram alterados. | [**CommentCollection**](/javascript/api/excel/excel.commentcollection#excel-excel-commentcollection-onchanged-member), [**Table**](/javascript/api/excel/excel.table#excel-excel-table-onchanged-member), [**TableCollection**](/javascript/api/excel/excel.tablecollection#excel-excel-tablecollection-onchanged-member), [**Worksheet**](/javascript/api/excel/excel.worksheet#excel-excel-worksheet-onchanged-member), [**WorksheetCollection**](/javascript/api/excel/excel.worksheetcollection#excel-excel-worksheetcollection-onchanged-member) |
| `onColumnSorted` | Ocorre quando uma ou mais colunas são classificadas. Isso acontece como resultado de uma operação de classificação da esquerda para a direita. | [**Planilha**](/javascript/api/excel/excel.worksheet#excel-excel-worksheet-oncolumnsorted-member), [**WorksheetCollection**](/javascript/api/excel/excel.worksheetcollection#excel-excel-worksheetcollection-oncolumnsorted-member) |
| `onDataChanged` | Ocorre quando os dados ou a formatação dentro da associação são alterados. | [**Associação**](/javascript/api/excel/excel.binding#excel-excel-binding-ondatachanged-member) |
| `onDeactivated` | Ocorre quando um objeto é desativado. | [**Gráfico**](/javascript/api/excel/excel.chart#excel-excel-chart-ondeactivated-member), [**ChartCollection**](/javascript/api/excel/excel.chartcollection#excel-excel-chartcollection-ondeactivated-member), [**Shape**](/javascript/api/excel/excel.shape#excel-excel-shape-ondeactivated-member), [**Planilha**](/javascript/api/excel/excel.worksheet#excel-excel-worksheet-ondeactivated-member), [**WorksheetCollection**](/javascript/api/excel/excel.worksheetcollection#excel-excel-worksheetcollection-ondeactivated-member) |
| `onDeleted` | Ocorre quando um objeto é excluído da coleção. | [**ChartCollection**](/javascript/api/excel/excel.chartcollection#excel-excel-chartcollection-ondeleted-member), [**CommentCollection**](/javascript/api/excel/excel.commentcollection#excel-excel-commentcollection-ondeleted-member), [**TableCollection**](/javascript/api/excel/excel.tablecollection#excel-excel-tablecollection-ondeleted-member), [**WorksheetCollection**](/javascript/api/excel/excel.worksheetcollection#excel-excel-worksheetcollection-ondeleted-member) |
| `onFormatChanged` | Ocorre quando o formato é alterado em uma planilha. | [**Planilha**](/javascript/api/excel/excel.worksheet#excel-excel-worksheet-onformatchanged-member), [**WorksheetCollection**](/javascript/api/excel/excel.worksheetcollection#excel-excel-worksheetcollection-onformatchanged-member) |
| `onFormulaChanged` | Ocorre quando uma fórmula é alterada. | [**Planilha**](/javascript/api/excel/excel.worksheet#excel-excel-worksheet-onformulachanged-member), [**WorksheetCollection**](/javascript/api/excel/excel.worksheetcollection#excel-excel-worksheetcollection-onformulachanged-member) |
| `onProtectionChanged` | Ocorre quando o estado de proteção da planilha é alterado. | [**Planilha**](/javascript/api/excel/excel.worksheet#excel-excel-worksheet-onprotectionchanged-member), [**WorksheetCollection**](/javascript/api/excel/excel.worksheetcollection#excel-excel-worksheetcollection-onprotectionchanged-member) |
| `onRowHiddenChanged` | Ocorre quando o estado de linha oculta é alterado em uma planilha específica. | [**Planilha**](/javascript/api/excel/excel.worksheet#excel-excel-worksheet-onrowhiddenchanged-member), [**WorksheetCollection**](/javascript/api/excel/excel.worksheetcollection#excel-excel-worksheetcollection-onrowhiddenchanged-member) |
| `onRowSorted` | Ocorre quando uma ou mais linhas são classificadas. Isso ocorre como resultado de uma operação de classificação de cima para baixo. | [**Planilha**](/javascript/api/excel/excel.worksheet#excel-excel-worksheet-onrowsorted-member), [**WorksheetCollection**](/javascript/api/excel/excel.worksheetcollection#excel-excel-worksheetcollection-onrowsorted-member) |
| `onSelectionChanged` | Ocorre quando uma célula ativa ou um intervalo selecionado são alterados. | [**Binding**](/javascript/api/excel/excel.binding#excel-excel-binding-onselectionchanged-member), [**Table**](/javascript/api/excel/excel.table#excel-excel-table-onselectionchanged-member), [**Workbook**](/javascript/api/excel/excel.workbook#excel-excel-workbook-onselectionchanged-member), [**Worksheet**](/javascript/api/excel/excel.worksheet#excel-excel-worksheet-onselectionchanged-member), [**WorksheetCollection**](/javascript/api/excel/excel.worksheetcollection#excel-excel-worksheetcollection-onselectionchanged-member) |
| `onSettingsChanged` | Ocorre quando as Configurações no documento são alteradas. | [**SettingCollection**](/javascript/api/excel/excel.settingcollection#excel-excel-settingcollection-onsettingschanged-member) |
| `onSingleClicked` | Acontece quando a operação é clicada/pressionada com o botão esquerdo do mouse ocorre na planilha. | [**Planilha**](/javascript/api/excel/excel.worksheet#excel-excel-worksheet-onsingleclicked-member), [**WorksheetCollection**](/javascript/api/excel/excel.worksheetcollection#excel-excel-worksheetcollection-onsingleclicked-member) |

### <a name="events-in-preview"></a>Eventos no modo de visualização

> [!NOTE]
> Os seguintes eventos estão disponíveis atualmente apenas na visualização pública. [!INCLUDE [Information about using preview APIs](../includes/using-excel-preview-apis.md)]

| Evento | Descrição | Objetos com suporte |
|:---------------|:-------------|:-----------|
| `onFiltered` | Ocorre quando um filtro é aplicado a um objeto. | [**Tabela**](/javascript/api/excel/excel.table#excel-excel-table-onfiltered-member), [**TableCollection**](/javascript/api/excel/excel.tablecollection#excel-excel-tablecollection-onfiltered-member), [**Planilha**](/javascript/api/excel/excel.worksheet#excel-excel-worksheet-onfiltered-member), [**WorksheetCollection**](/javascript/api/excel/excel.worksheetcollection#excel-excel-worksheetcollection-onfiltered-member) |

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
await Excel.run(async (context) => {
    const worksheet = context.workbook.worksheets.getItem("Sample");
    worksheet.onChanged.add(handleChange);

    await context.sync();
    console.log("Event handler successfully registered for onChanged event in the worksheet.");
}).catch(errorHandlerFunction);
```

## <a name="handle-an-event"></a>Manipular um evento

Como mostrado no exemplo anterior, quando você registrar um manipulador de eventos, indica a função a ser executada quando o evento especificado ocorre. Você pode criar essa função para executar as ações que seu cenário exige. O exemplo de código a seguir mostra uma função de manipulador de eventos que simplesmente grava informações sobre o evento no console.

```js
async function handleChange(event) {
    await Excel.run(async (context) => {
        await context.sync();        
        console.log("Change type of event: " + event.changeType);
        console.log("Address of event: " + event.address);
        console.log("Source of event: " + event.source);       
    }).catch(errorHandlerFunction);
}
```

## <a name="remove-an-event-handler"></a>Remover um manipulador de eventos

O exemplo de código a seguir registra um manipulador de eventos para o evento `onSelectionChanged` na planilha **Sample** e define a função `handleSelectionChange` a executar quando o evento ocorrer. Também define a função `remove()` que pode ser chamada posteriormente para remover aquele manipulador de eventos. Observe que o usado `RequestContext` para criar o manipulador de eventos é necessário para removê-lo.

```js
let eventResult;

async function run() {
  await Excel.run(async (context) => {
    const worksheet = context.workbook.worksheets.getItem("Sample");
    eventResult = worksheet.onSelectionChanged.add(handleSelectionChange);

    await context.sync();
    console.log("Event handler successfully registered for onSelectionChanged event in the worksheet.");
  });
}

async function handleSelectionChange(event) {
  await Excel.run(async (context) => {
    await context.sync();
    console.log("Address of current selection: " + event.address);
  });
}

async function remove() {
  await Excel.run(eventResult.context, async (context) => {
    eventResult.remove();
    await context.sync();
    
    eventResult = null;
    console.log("Event handler successfully removed.");
  });
}
```

## <a name="enable-and-disable-events"></a>Habilitar e desabilitar eventos

O desempenho de um suplemento pode ser melhorado desabilitando eventos.
Por exemplo, seu aplicativo pode não precisar receber eventos ou ele pode ignorar eventos durante a edições de lotes de várias entidades.

Os eventos são habilitados ou desabilitados no nível [runtime](/javascript/api/excel/excel.runtime).
A `enableEvents` propriedade determina se os eventos são disparados e se seus manipuladores são ativados.

O código a seguir mostra como ativar ou desativar os eventos.

```js
await Excel.run(async (context) => {
    context.runtime.load("enableEvents");
    await context.sync();

    let eventBoolean = !context.runtime.enableEvents;
    context.runtime.enableEvents = eventBoolean;
    if (eventBoolean) {
        console.log("Events are currently on.");
    } else {
        console.log("Events are currently off.");
    }
    
    await context.sync();
});
```

## <a name="see-also"></a>Confira também

- [Modelo de objeto JavaScript do Excel em Suplementos do Office](excel-add-ins-core-concepts.md)
