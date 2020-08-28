---
title: Otimização de desempenho do da API JavaScript do Excel
description: Otimizar o desempenho do suplemento do Excel usando a API JavaScript.
ms.date: 07/29/2020
localization_priority: Normal
ms.openlocfilehash: fdaccdca4779aaca64420794e382330994488606
ms.sourcegitcommit: 9609bd5b4982cdaa2ea7637709a78a45835ffb19
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 08/28/2020
ms.locfileid: "47294098"
---
# <a name="performance-optimization-using-the-excel-javascript-api"></a>Otimização de desempenho usando a API JavaScript do Excel

Existem várias maneiras de executar tarefas comuns com a API JavaScript do Excel. Você encontrará diferenças significativas de desempenho entre várias abordagens. Este artigo fornece orientações e amostras de código para mostrar como realizar tarefas comuns com eficiência usando as API JavaScript do Excel.

> [!IMPORTANT]
> Muitos problemas de desempenho podem ser tratados através do uso recomendado de `load` `sync` chamadas e. Consulte a seção "aprimoramentos de desempenho com as APIs específicas do aplicativo" de [limites de recurso e otimização de desempenho para suplementos do Office](../concepts/resource-limits-and-performance-optimization.md#performance-improvements-with-the-application-specific-apis) para conselhos sobre como trabalhar com APIs específicas do aplicativo de uma maneira eficiente.

## <a name="suspend-excel-processes-temporarily"></a>Suspender temporariamente os processos do Excel

O Excel tem várias tarefas em segundo plano reagindo à entrada de usuários e seu suplemento. Alguns desses processos do Excel podem ser controlado para obter o benefício de desempenho. Isso é útil principalmente quando o suplemento lida com grandes conjuntos de dados.

### <a name="suspend-calculation-temporarily"></a>Suspender os cálculos temporariamente

Se você estiver tentando executar uma operação em um grande número de células (por exemplo, definindo o valor do objeto de um grande intervalo) e não se importar em suspender o cálculo no Excel temporariamente enquanto a operação for concluída, é recomendável que você suspenda o cálculo até o próximo `context.sync()` ser chamado.

Ver a documentação de referência [objeto de aplicativo](/javascript/api/excel/excel.application) para saber mais sobre como usar a API`suspendApiCalculationUntilNextSync()`para suspender e reativar cálculos de maneira muito fácil. O código a seguir demonstra como suspender temporariamente um cálculo:

```js
Excel.run(async function(ctx) {
    var app = ctx.workbook.application;
    var sheet = ctx.workbook.worksheets.getItem("sheet1");
    var rangeToSet: Excel.Range;
    var rangeToGet: Excel.Range;
    app.load("calculationMode");
    await ctx.sync();
    // Calculation mode should be "Automatic" by default
    console.log(app.calculationMode);

    rangeToSet = sheet.getRange("A1:C1");
    rangeToSet.values = [[1, 2, "=SUM(A1:B1)"]];
    rangeToGet = sheet.getRange("A1:C1");
    rangeToGet.load("values");
    await ctx.sync();
    // Range value should be [1, 2, 3] now
    console.log(rangeToGet.values);

    // Suspending recalculation
    app.suspendApiCalculationUntilNextSync();
    rangeToSet = sheet.getRange("A1:B1");
    rangeToSet.values = [[10, 20]];
    rangeToGet = sheet.getRange("A1:C1");
    rangeToGet.load("values");
    app.load("calculationMode");
    await ctx.sync();
    // Range value should be [10, 20, 3] when we load the property, because calculation is suspended at that point
    console.log(rangeToGet.values);
    // Calculation mode should still be "Automatic" even with suspend recalculation
    console.log(app.calculationMode);

    rangeToGet.load("values");
    await ctx.sync();
    // Range value should be [10, 20, 30] when we load the property, because calculation is resumed after last sync
    console.log(rangeToGet.values);
})
```

Observe que somente os cálculos de fórmula são suspensos. Todas as referências alteradas ainda serão recriadas. Por exemplo, renomear uma planilha ainda atualiza quaisquer referências em fórmulas para essa planilha.

### <a name="suspend-screen-updating"></a>Suspender a atualização da tela

O Excel exibe as alterações que seu suplemento faz aproximadamente conforme elas acontecem no código. Para conjuntos de dados grandes e interativos, talvez não seja necessário não esse andamento na tela em tempo real. `Application.suspendScreenUpdatingUntilNextSync()` pausa atualizações visuais no Excel até as chamadas do suplemento `context.sync()`, ou até o`Excel.run` terminar (chamadas implícitas `context.sync`). Lembre-se, o Excel não mostrará os sinais de atividade até a próxima sincronização. Seu suplemento deve fornecer orientação aos usuários para prepará-los para esse atraso ou fornecer uma barra de status para demonstrar atividade.

> [!NOTE]
> Não chame `suspendScreenUpdatingUntilNextSync` repetidamente (como em um loop). As chamadas repetidas farão com que a janela do Excel fique de piscar.

### <a name="enable-and-disable-events"></a>Habilitar e desabilitar eventos

O desempenho de um suplemento pode ser melhorado desabilitando eventos. Um exemplo de código mostrando como habilitar e desabilitar os eventos está no artigo [trabalhar com eventos](excel-add-ins-events.md#enable-and-disable-events).

## <a name="importing-data-into-tables"></a>Importar dados em tabelas

Ao tentar importar um grande volume de dados diretamente em um objeto[tabela](/javascript/api/excel/excel.table) diretamente (por exemplo, usando `TableRowCollection.add()`), você poderá observar um desempenho lento. Se você estiver tentando adicionar uma nova tabela, você deve preencher os dados primeiro definindo `range.values`e em seguida, ligue `worksheet.tables.add()` para criar uma tabela de intervalo. Se você está tentando gravar dados em uma tabela existente, grave os dados em um intervalo de objeto via`table.getDataBodyRange()`, e a tabela será expandida automaticamente.

Aqui está um exemplo dessa abordagem:

```js
Excel.run(async (ctx) => {
    var sheet = ctx.workbook.worksheets.getItem("Sheet1");
    // Write the data into the range first.
    var range = sheet.getRange("A1:B3");
    range.values = [["Key", "Value"], ["A", 1], ["B", 2]];

    // Create the table over the range
    var table = sheet.tables.add('A1:B3', true);
    table.name = "Example";
    await ctx.sync();


    // Insert a new row to the table
    table.getDataBodyRange().getRowsBelow(1).values = [["C", 3]];
    // Change a existing row value
    table.getDataBodyRange().getRow(1).values = [["D", 4]];
    await ctx.sync();
})
```

> [!NOTE]
> Você pode converter convenientemente um objeto de tabela em um objeto de intervalo usando o método[Table.convertToRange()](/javascript/api/excel/excel.table#converttorange--).

## <a name="see-also"></a>Confira também

* [Conceitos fundamentais de programação com a API JavaScript do Excel](excel-add-ins-core-concepts.md)
* [Limites de recurso e otimização de desempenho para Suplementos do Office](../concepts/resource-limits-and-performance-optimization.md)
* [Objeto de funções de planilha (API JavaScript para Excel)](/javascript/api/excel/excel.functions)
