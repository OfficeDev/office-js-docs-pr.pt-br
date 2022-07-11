---
title: Otimização de desempenho do da API JavaScript do Excel
description: Otimize o desempenho do suplemento do Excel usando a API JavaScript.
ms.date: 02/17/2022
ms.localizationpriority: medium
ms.openlocfilehash: bad5d35ec1cc3f99cd37b3571dee78d3432102e6
ms.sourcegitcommit: d8ea4b761f44d3227b7f2c73e52f0d2233bf22e2
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 07/11/2022
ms.locfileid: "66712724"
---
# <a name="performance-optimization-using-the-excel-javascript-api"></a>Otimização de desempenho usando a API JavaScript do Excel

Existem várias maneiras de executar tarefas comuns com a API JavaScript do Excel. Você encontrará diferenças significativas de desempenho entre várias abordagens. Este artigo fornece orientações e amostras de código para mostrar como realizar tarefas comuns com eficiência usando as API JavaScript do Excel.

> [!IMPORTANT]
> Muitos problemas de desempenho podem ser resolvidos por meio do uso recomendado de `load` chamadas `sync` . Consulte a seção "Melhorias de desempenho com as APIs específicas do aplicativo" dos limites de recursos e otimização de desempenho para [Suplementos do Office](../concepts/resource-limits-and-performance-optimization.md#performance-improvements-with-the-application-specific-apis) para obter conselhos sobre como trabalhar com as APIs específicas do aplicativo de maneira eficiente.

## <a name="suspend-excel-processes-temporarily"></a>Suspender temporariamente os processos do Excel

O Excel tem várias tarefas em segundo plano reagindo à entrada de usuários e seu suplemento. Alguns desses processos do Excel podem ser controlado para obter o benefício de desempenho. Isso é útil principalmente quando o suplemento lida com grandes conjuntos de dados.

### <a name="suspend-calculation-temporarily"></a>Suspender os cálculos temporariamente

Se você estiver tentando executar uma operação em um grande número de células (por exemplo, definindo o valor do objeto de um grande intervalo) e não se importar em suspender o cálculo no Excel temporariamente enquanto a operação for concluída, é recomendável que você suspenda o cálculo até o próximo `context.sync()` ser chamado.

Ver a documentação de referência [objeto de aplicativo](/javascript/api/excel/excel.application) para saber mais sobre como usar a API`suspendApiCalculationUntilNextSync()`para suspender e reativar cálculos de maneira muito fácil. O código a seguir demonstra como suspender o cálculo temporariamente.

```js
await Excel.run(async (context) => {
    let app = context.workbook.application;
    let sheet = context.workbook.worksheets.getItem("sheet1");
    let rangeToSet: Excel.Range;
    let rangeToGet: Excel.Range;
    app.load("calculationMode");
    await context.sync();
    // Calculation mode should be "Automatic" by default
    console.log(app.calculationMode);

    rangeToSet = sheet.getRange("A1:C1");
    rangeToSet.values = [[1, 2, "=SUM(A1:B1)"]];
    rangeToGet = sheet.getRange("A1:C1");
    rangeToGet.load("values");
    await context.sync();
    // Range value should be [1, 2, 3] now
    console.log(rangeToGet.values);

    // Suspending recalculation
    app.suspendApiCalculationUntilNextSync();
    rangeToSet = sheet.getRange("A1:B1");
    rangeToSet.values = [[10, 20]];
    rangeToGet = sheet.getRange("A1:C1");
    rangeToGet.load("values");
    app.load("calculationMode");
    await context.sync();
    // Range value should be [10, 20, 3] when we load the property, because calculation is suspended at that point
    console.log(rangeToGet.values);
    // Calculation mode should still be "Automatic" even with suspend recalculation
    console.log(app.calculationMode);

    rangeToGet.load("values");
    await context.sync();
    // Range value should be [10, 20, 30] when we load the property, because calculation is resumed after last sync
    console.log(rangeToGet.values);
});
```

Observe que apenas os cálculos de fórmula são suspensos. Todas as referências alteradas ainda são recriadas. Por exemplo, renomear uma planilha ainda atualiza quaisquer referências em fórmulas para essa planilha.

### <a name="suspend-screen-updating"></a>Suspender a atualização da tela

O Excel exibe as alterações que seu suplemento faz aproximadamente conforme elas acontecem no código. Para conjuntos de dados grandes e interativos, talvez não seja necessário não esse andamento na tela em tempo real. `Application.suspendScreenUpdatingUntilNextSync()` pausa atualizações visuais no Excel até as chamadas do suplemento `context.sync()`, ou até o`Excel.run` terminar (chamadas implícitas `context.sync`). Lembre-se, o Excel não mostrará os sinais de atividade até a próxima sincronização. Seu suplemento deve fornecer orientação aos usuários para prepará-los para esse atraso ou fornecer uma barra de status para demonstrar atividade.

> [!NOTE]
> Não chame repetidamente `suspendScreenUpdatingUntilNextSync` (como em um loop). Chamadas repetidas farão a janela do Excel piscar.

### <a name="enable-and-disable-events"></a>Habilitar e desabilitar eventos

O desempenho de um suplemento pode ser melhorado desabilitando eventos. Um exemplo de código mostrando como habilitar e desabilitar os eventos está no artigo [trabalhar com eventos](excel-add-ins-events.md#enable-and-disable-events).

## <a name="importing-data-into-tables"></a>Importar dados em tabelas

Ao tentar importar um grande volume de dados diretamente em um objeto[tabela](/javascript/api/excel/excel.table) diretamente (por exemplo, usando `TableRowCollection.add()`), você poderá observar um desempenho lento. Se você estiver tentando adicionar uma nova tabela, você deve preencher os dados primeiro definindo `range.values`e em seguida, ligue `worksheet.tables.add()` para criar uma tabela de intervalo. Se você está tentando gravar dados em uma tabela existente, grave os dados em um intervalo de objeto via`table.getDataBodyRange()`, e a tabela será expandida automaticamente.

Aqui está um exemplo dessa abordagem:

```js
await Excel.run(async (context) => {
    let sheet = context.workbook.worksheets.getItem("Sheet1");
    // Write the data into the range first.
    let range = sheet.getRange("A1:B3");
    range.values = [["Key", "Value"], ["A", 1], ["B", 2]];

    // Create the table over the range
    let table = sheet.tables.add('A1:B3', true);
    table.name = "Example";
    await context.sync();


    // Insert a new row to the table
    table.getDataBodyRange().getRowsBelow(1).values = [["C", 3]];
    // Change a existing row value
    table.getDataBodyRange().getRow(1).values = [["D", 4]];
    await context.sync();
});
```

> [!NOTE]
> Você pode converter convenientemente um objeto de tabela em um objeto de intervalo usando o método[Table.convertToRange()](/javascript/api/excel/excel.table#excel-excel-table-converttorange-member(1)).

## <a name="payload-size-limit-best-practices"></a>Práticas recomendadas de limite de tamanho de conteúdo

A API JavaScript do Excel tem limitações de tamanho para chamadas à API. Excel na Web tem um limite de tamanho de carga para solicitações e respostas de 5 MB, e uma API `RichAPI.Error` retornará um erro se esse limite for excedido. Em todas as plataformas, um intervalo é limitado a cinco milhões de células para obter operações. Intervalos grandes geralmente excedem essas duas limitações.

O tamanho da carga de uma solicitação é uma combinação dos três componentes a seguir.

* O número de chamadas à API
* O número de objetos, como `Range` objetos
* O comprimento do valor a ser definido ou obtido

Se uma API retornar o `RequestPayloadSizeLimitExceeded` erro, use as estratégias de melhores práticas documentadas neste artigo para otimizar o script e evitar o erro.

### <a name="strategy-1-move-unchanged-values-out-of-loops"></a>Estratégia 1: mover valores inalterados de loops

Limite o número de processos que ocorrem em loops para melhorar o desempenho. No exemplo de código a seguir, `context.workbook.worksheets.getActiveWorksheet()` pode ser movido para fora `for` do loop, porque ele não é alterado dentro desse loop.

```js
// DO NOT USE THIS CODE SAMPLE. This sample shows a poor performance strategy. 
async function run() {
  await Excel.run(async (context) => {
    let ranges = [];
    
    // This sample retrieves the worksheet every time the loop runs, which is bad for performance.
    for (let i = 0; i < 7500; i++) {
      let rangeByIndex = context.workbook.worksheets.getActiveWorksheet().getRangeByIndexes(i, 1, 1, 1);
    }    
    await context.sync();
  });
}
```

O exemplo de código a seguir mostra uma lógica semelhante ao exemplo de código anterior, mas com uma estratégia de desempenho aprimorada. O valor `context.workbook.worksheets.getActiveWorksheet()` é recuperado antes do `for` loop, porque esse valor não precisa ser recuperado sempre que o loop é `for` executado. Somente valores que mudam dentro do contexto de um loop devem ser recuperados dentro desse loop.

```js
// This code sample shows a good performance strategy.
async function run() {
  await Excel.run(async (context) => {
    let ranges = [];
    // Retrieve the worksheet outside the loop.
    let worksheet = context.workbook.worksheets.getActiveWorksheet(); 

    // Only process the necessary values inside the loop.
    for (let i = 0; i < 7500; i++) {
      let rangeByIndex = worksheet.getRangeByIndexes(i, 1, 1, 1);
    }    
    await context.sync();
  });
}
```

### <a name="strategy-2-create-fewer-range-objects"></a>Estratégia 2: criar menos objetos de intervalo

Crie menos objetos de intervalo para melhorar o desempenho e minimizar o tamanho da carga. Duas abordagens para criar menos objetos de intervalo são descritas nas seções e exemplos de código do artigo a seguir.

#### <a name="split-each-range-array-into-multiple-arrays"></a>Dividir cada matriz de intervalo em várias matrizes

Uma maneira de criar menos objetos de intervalo é dividir cada matriz de intervalo em várias matrizes e, em seguida, processar cada nova matriz com um loop e uma nova `context.sync()` chamada.

> [!IMPORTANT]
> Use essa estratégia somente se você tiver determinado primeiro que está excedendo o limite de tamanho da solicitação de conteúdo. O uso de vários loops pode reduzir o tamanho de cada solicitação de carga para evitar exceder o limite de 5 MB, mas o uso de vários loops `context.sync()` e várias chamadas também afeta negativamente o desempenho.

O exemplo de código a seguir tenta processar uma grande matriz de intervalos em um único loop e, em seguida, uma única `context.sync()` chamada. O processamento de muitos valores de intervalo em uma `context.sync()` chamada faz com que o tamanho da solicitação de conteúdo exceda o limite de 5 MB.

```js
// This code sample does not show a recommended strategy.
// Calling 10,000 rows would likely exceed the 5MB payload size limit in a real-world situation.
async function run() {
  await Excel.run(async (context) => {
    let worksheet = context.workbook.worksheets.getActiveWorksheet();
    
    // This sample attempts to process too many ranges at once. 
    for (let row = 1; row < 10000; row++) {
      let range = sheet.getRangeByIndexes(row, 1, 1, 1);
      range.values = [["1"]];
    }
    await context.sync(); 
  });
}
```

O exemplo de código a seguir mostra uma lógica semelhante ao exemplo de código anterior, mas com uma estratégia que evita exceder o limite de tamanho da solicitação de carga de 5 MB. No exemplo de código a seguir, os intervalos são processados em dois loops separados e cada loop é seguido por uma `context.sync()` chamada.

```js
// This code sample shows a strategy for reducing payload request size.
// However, using multiple loops and `context.sync()` calls negatively impacts performance.
// Only use this strategy if you've determined that you're exceeding the payload request limit.
async function run() {
  await Excel.run(async (context) => {
    let worksheet = context.workbook.worksheets.getActiveWorksheet();

    // Split the ranges into two loops, rows 1-5000 and then 5001-10000.
    for (let row = 1; row < 5000; row++) {
      let range = worksheet.getRangeByIndexes(row, 1, 1, 1);
      range.values = [["1"]];
    }
    // Sync after each loop. 
    await context.sync(); 
    
    for (let row = 5001; row < 10000; row++) {
      let range = worksheet.getRangeByIndexes(row, 1, 1, 1);
      range.values = [["1"]];
    }
    await context.sync(); 
  });
}
```

#### <a name="set-range-values-in-an-array"></a>Definir valores de intervalo em uma matriz

Outra maneira de criar menos objetos de intervalo é criar uma matriz, usar um loop para definir todos os dados nessa matriz e, em seguida, passar os valores da matriz para um intervalo. Isso beneficia o desempenho e o tamanho da carga. Em vez de chamar `range.values` cada intervalo em um loop, `range.values` é chamado uma vez fora do loop.

O exemplo de código a `for` seguir mostra como criar uma matriz, definir os valores dessa matriz em um loop e, em seguida, passar os valores de matriz para um intervalo fora do loop.

```js
// This code sample shows a good performance strategy.
async function run() {
  await Excel.run(async (context) => {
    const worksheet = context.workbook.worksheets.getActiveWorksheet();    
    // Create an array.
    const array = new Array(10000);

    // Set the values of the array inside the loop.
    for (let i = 0; i < 10000; i++) {
      array[i] = [1];
    }

    // Pass the array values to a range outside the loop. 
    let range = worksheet.getRange("A1:A10000");
    range.values = array;
    await context.sync();
  });
}
```

## <a name="see-also"></a>Confira também

* [Modelo de objeto JavaScript do Excel em Suplementos do Office](excel-add-ins-core-concepts.md)
* [Tratamento de erro com as APIs JavaScript específicas do aplicativo](../testing/application-specific-api-error-handling.md)
* [Limites de recurso e otimização de desempenho para Suplementos do Office](../concepts/resource-limits-and-performance-optimization.md)
* [Objeto de funções de planilha (API JavaScript para Excel)](/javascript/api/excel/excel.functions)
