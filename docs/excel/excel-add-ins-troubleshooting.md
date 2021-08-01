---
title: Solução de Excel de solução de problemas
description: Saiba como solucionar erros de desenvolvimento em Excel de complementos.
ms.date: 02/12/2021
localization_priority: Normal
ms.openlocfilehash: b90d8cfdb4696445655122a2fa7eb74d1c87fa2f
ms.sourcegitcommit: 3fa8c754a47bab909e559ae3e5d4237ba27fdbe4
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 07/30/2021
ms.locfileid: "53671461"
---
# <a name="troubleshooting-excel-add-ins"></a>Solução de Excel de solução de problemas

Este artigo discute a solução de problemas que são exclusivos Excel. Use a ferramenta de comentários na parte inferior da página para sugerir outros problemas que podem ser adicionados ao artigo.

## <a name="api-limitations-when-the-active-workbook-switches"></a>Limitações da API quando a agenda de trabalho ativa é alternada

Os complementos para Excel são destinados a operar em uma única workbook de cada vez. Erros podem surgir quando uma workbook separada da que está executando o complemento ganha o foco. Isso só acontece quando determinados métodos estão no processo de ser chamado quando o foco muda.

As APIs a seguir são afetadas por essa opção de lista de trabalho.

|API JavaScript do Excel | Erro lançado |
|--|--|
| `Chart.activate` | GeneralException |
| `Range.select` | GeneralException |
| `Table.clearFilters` | GeneralException |
| `Workbook.getActiveCell`  | InvalidSelection|
| `Workbook.getSelectedRange` | InvalidSelection|
| `Workbook.getSelectedRanges`  | InvalidSelection|
| `Worksheet.activate` | GeneralException |
| `Worksheet.delete`  | InvalidSelection|
| `Worksheet.gridlines` | GeneralException |
| `Worksheet.showHeadings` | GeneralException |
| `WorksheetCollection.add` | GeneralException |
| `WorksheetFreezePanes.freezeAt` | GeneralException |
| `WorksheetFreezePanes.freezeColumns` | GeneralException |
| `WorksheetFreezePanes.freezeRows` | GeneralException |
| `WorksheetFreezePanes.getLocationOrNullObject`| GeneralException |
| `WorksheetFreezePanes.unfreeze` | GeneralException |

> [!NOTE]
> Isso só se aplica a várias Excel de trabalho abertas em Windows ou Mac.

## <a name="coauthoring"></a>Coautoria

Consulte [Coautor no Excel para](co-authoring-in-excel-add-ins.md) padrões a ser usado com eventos em um ambiente de coautor. O artigo também aborda possíveis conflitos de mesclagem ao usar determinadas APIs, como [`TableRowCollection.add`](/javascript/api/excel/excel.tablerowcollection#add_index__values_) .

## <a name="known-issues"></a>Problemas Conhecidos

### <a name="binding-events-return-temporary-binding-obects"></a>Eventos de associação `Binding` retornam obects temporários

[BindingDataChangedEventArgs.binding](/javascript/api/excel/excel.bindingdatachangedeventargs#binding) e [BindingSelectionChangedEventArgs.binding](/javascript/api/excel/excel.bindingselectionchangedeventargs#binding) retornam um objeto temporário que contém a ID do objeto que gerou o `Binding` `Binding` evento. Use essa ID com `BindingCollection.getItem(id)` para recuperar o objeto que gerou o `Binding` evento.

O exemplo de código a seguir mostra como usar essa ID de associação temporária para recuperar o objeto `Binding` relacionado. No exemplo, um ouvinte de eventos é atribuído a uma associação. O ouvinte chama `getBindingId` o método quando o evento é `onDataChanged` disparado. O `getBindingId` método usa a ID do objeto temporário para recuperar o objeto que gerou o `Binding` `Binding` evento.

```js
Excel.run(function (context) {
    // Retrieve your binding.
    var binding = context.workbook.bindings.getItemAt(0);

    return context.sync().then(function () {
        // Register an event listener to detect changes to your binding
        // and then trigger the `getBindingId` method when the data changes. 
        binding.onDataChanged.add(getBindingId);

        return context.sync();
    });
});

function getBindingId(eventArgs) {
    return Excel.run(function (context) {
        // Get the temporary binding object and load its ID. 
        var tempBindingObject = eventArgs.binding;
        tempBindingObject.load("id");

        // Use the temporary binding object's ID to retrieve the original binding object. 
        var originalBindingObject = context.workbook.bindings.getItem(tempBindingObject.id);

        // You now have the binding object that raised the event: `originalBindingObject`. 
    });
}
```

### <a name="cell-format-usestandardheight-and-usestandardwidth-issues"></a>Formato de célula `useStandardHeight` e `useStandardWidth` problemas

A [propriedade useStandardHeight](/javascript/api/excel/excel.cellpropertiesformat#useStandardHeight) `CellPropertiesFormat` não funciona corretamente no Excel na Web. Devido a um problema na interface do usuário Excel na Web, definir a propriedade para calcular a altura `useStandardHeight` `true` imprecisamente nessa plataforma. Por exemplo, uma altura padrão **de 14** é modificada para **14,25** em Excel na Web.

Em todas as plataformas, [as propriedades useStandardHeight](/javascript/api/excel/excel.cellpropertiesformat#useStandardHeight) e [useStandardWidth](/javascript/api/excel/excel.cellpropertiesformat#useStandardWidth) devem ser definidas `CellPropertiesFormat` apenas como `true` . Definir essas propriedades como `false` não tem efeito. 

### <a name="range-getimage-method-unsupported-on-excel-for-mac"></a>Método `getImage` Range sem suporte no Excel para Mac

O método [Range getImage](/javascript/api/excel/excel.range#getImage__) não tem suporte no Excel para Mac. Consulte [OfficeDev/office-js Issue #235](https://github.com/OfficeDev/office-js/issues/235) para o status atual.

### <a name="range-return-character-limit"></a>Limite de caracteres de retorno de intervalo

Os [métodos Worksheet.getRange(address)](/javascript/api/excel/excel.worksheet#getRange_address_) e [Worksheet.getRanges(address)](/javascript/api/excel/excel.worksheet#getRanges_address_) têm um limite de cadeia de caracteres de endereço de 8192 caracteres. Quando esse limite é excedido, a cadeia de caracteres de endereço é truncada para 8192 caracteres.

## <a name="see-also"></a>Conferir também

- [Solucionar erros de desenvolvimento com Office de complementos](../testing/troubleshoot-development-errors.md)
- [Solucionar erros de usuários com Suplementos do Office](../testing/testing-and-troubleshooting.md)
