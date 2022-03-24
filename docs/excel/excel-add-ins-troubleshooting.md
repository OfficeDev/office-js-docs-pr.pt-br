---
title: Solução de Excel de soluções de problemas
description: Saiba como solucionar erros de desenvolvimento em Excel de complementos.
ms.date: 02/17/2022
ms.localizationpriority: medium
ms.openlocfilehash: aabbb3d8b62101eacb2ac51684a3d1f6c16e84a4
ms.sourcegitcommit: 968d637defe816449a797aefd930872229214898
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 03/23/2022
ms.locfileid: "63745509"
---
# <a name="troubleshooting-excel-add-ins"></a>Solução de Excel de soluções de problemas

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

Consulte [Coautor no Excel para](co-authoring-in-excel-add-ins.md) padrões a ser usado com eventos em um ambiente de coautor. O artigo também aborda possíveis conflitos de mesclagem ao usar determinadas APIs, como [`TableRowCollection.add`](/javascript/api/excel/excel.tablerowcollection#excel-excel-tablerowcollection-add-member(1)).

## <a name="known-issues"></a>Problemas Conhecidos

### <a name="binding-events-return-temporary-binding-obects"></a>Eventos de associação retornam `Binding` obects temporários

[BindingDataChangedEventArgs.binding](/javascript/api/excel/excel.bindingdatachangedeventargs#excel-excel-bindingdatachangedeventargs-binding-member) e [BindingSelectionChangedEventArgs.binding](/javascript/api/excel/excel.bindingselectionchangedeventargs#excel-excel-bindingselectionchangedeventargs-binding-member) retornam um objeto temporário que contém a `Binding` ID `Binding` do objeto que gerou o evento. Use essa ID com `BindingCollection.getItem(id)` para recuperar o `Binding` objeto que gerou o evento.

O exemplo de código a seguir mostra como usar essa ID de associação temporária para recuperar o objeto `Binding` relacionado. No exemplo, um ouvinte de eventos é atribuído a uma associação. O ouvinte chama o `getBindingId` método quando o `onDataChanged` evento é disparado. O `getBindingId` método usa a ID do objeto temporário `Binding` para recuperar o `Binding` objeto que gerou o evento.

```js
async function run() {
    await Excel.run(async (context) => {
        // Retrieve your binding.
        let binding = context.workbook.bindings.getItemAt(0);
    
        await context.sync();
    
        // Register an event listener to detect changes to your binding
        // and then trigger the `getBindingId` method when the data changes. 
        binding.onDataChanged.add(getBindingId);
        await context.sync();
    });
}

async function getBindingId(eventArgs) {
    await Excel.run(async (context) => {
        // Get the temporary binding object and load its ID. 
        let tempBindingObject = eventArgs.binding;
        tempBindingObject.load("id");

        // Use the temporary binding object's ID to retrieve the original binding object. 
        let originalBindingObject = context.workbook.bindings.getItem(tempBindingObject.id);

        // You now have the binding object that raised the event: `originalBindingObject`. 
    });
}
```

### <a name="cell-format-usestandardheight-and-usestandardwidth-issues"></a>Formato de célula e `useStandardHeight` `useStandardWidth` problemas

A [propriedade useStandardHeight](/javascript/api/excel/excel.cellpropertiesformat#excel-excel-cellpropertiesformat-usestandardheight-member) não `CellPropertiesFormat` funciona corretamente no Excel na Web. Devido a um problema na interface do usuário Excel na Web, `useStandardHeight` `true` definir a propriedade para calcular a altura de forma impreciso nessa plataforma. Por exemplo, uma altura padrão **de 14** é modificada para **14,25** em Excel na Web.

Em todas as plataformas, [as propriedades useStandardHeight](/javascript/api/excel/excel.cellpropertiesformat#excel-excel-cellpropertiesformat-usestandardheight-member) e [useStandardWidth](/javascript/api/excel/excel.cellpropertiesformat#excel-excel-cellpropertiesformat-usestandardwidth-member) `CellPropertiesFormat` devem ser definidas apenas como `true`. Definir essas propriedades como não `false` tem efeito.

### <a name="range-getimage-method-unsupported-on-excel-for-mac"></a>Método Range `getImage` sem suporte no Excel para Mac

O método [Range getImage](/javascript/api/excel/excel.range#excel-excel-range-getimage-member(1)) não tem suporte no momento Excel para Mac. Consulte [OfficeDev/office-js Problema #235](https://github.com/OfficeDev/office-js/issues/235) para o status atual.

### <a name="range-return-character-limit"></a>Limite de caracteres de retorno de intervalo

Os [métodos Worksheet.getRange(address)](/javascript/api/excel/excel.worksheet#excel-excel-worksheet-getrange-member(1)) e [Worksheet.getRanges(address)](/javascript/api/excel/excel.worksheet#excel-excel-worksheet-getranges-member(1)) têm um limite de cadeia de caracteres de endereço de 8192 caracteres. Quando esse limite é excedido, a cadeia de caracteres de endereço é truncada para 8192 caracteres.

## <a name="see-also"></a>Veja também

- [Solucionar erros de desenvolvimento com Suplementos do Office](../testing/troubleshoot-development-errors.md)
- [Solucionar erros de usuários com Suplementos do Office](../testing/testing-and-troubleshooting.md)
