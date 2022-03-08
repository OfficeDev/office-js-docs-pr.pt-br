---
title: Trabalhar com tabelas dinâmicas usando a Excel JavaScript
description: Use a Excel JavaScript para criar tabelas dinâmicas e interagir com seus componentes.
ms.date: 02/16/2022
ms.localizationpriority: medium
ms.openlocfilehash: 5a47baf51a371a388959acbc56778e04f72bcd57
ms.sourcegitcommit: 7b6ee73fa70b8e0ff45c68675dd26dd7a7b8c3e9
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 03/08/2022
ms.locfileid: "63340369"
---
# <a name="work-with-pivottables-using-the-excel-javascript-api"></a>Trabalhar com tabelas dinâmicas usando a Excel JavaScript

Tabelas dinâmicas simplificam conjuntos de dados maiores. Eles permitem a manipulação rápida de dados agrupados. A Excel api JavaScript permite que seu complemento crie Tabelas Dinâmicas e interaja com seus componentes. Este artigo descreve como as Tabelas Dinâmicas são representadas pela API JavaScript Office e fornece exemplos de código para cenários principais.

Se você não estiver familiarizado com a funcionalidade das Tabelas Dinâmicas, considere explorá-las como um usuário final.
Consulte [Criar uma Tabela Dinâmica para analisar dados de planilha](https://support.microsoft.com/office/ccd3c4a6-272f-4c97-afbb-d3f27407fcde#ID0EBBD=PivotTables) para uma boa cartilha nessas ferramentas.

> [!IMPORTANT]
> As Tabelas Dinâmicas criadas com o OLAP não são suportadas no momento. Também não há suporte para o Power Pivot.

## <a name="object-model"></a>Modelo de objetos

A [Tabela Dinâmica](/javascript/api/excel/excel.pivottable) é o objeto central para Tabelas Dinâmicas na API JavaScript Office JavaScript.

- `Workbook.pivotTables` e `Worksheet.pivotTables` são [PivotTableCollections](/javascript/api/excel/excel.pivottablecollection) que contêm as [Tabelas Dinâmicas](/javascript/api/excel/excel.pivottable) na pasta de trabalho e planilha, respectivamente.
- Uma [Tabela Dinâmica](/javascript/api/excel/excel.pivottable) contém [um PivotHierarchyCollection](/javascript/api/excel/excel.pivothierarchycollection) que tem vários [PivotHierarchies](/javascript/api/excel/excel.pivothierarchy).
- Essas [PivotHierarchies](/javascript/api/excel/excel.pivothierarchy) podem ser adicionadas a coleções de hierarquia específicas para definir como os dados de pivot de tabela dinâmica (conforme explicado na [seção a seguir](#hierarchies)).
- Um [PivotHierarchy](/javascript/api/excel/excel.pivothierarchy) contém [um PivotFieldCollection](/javascript/api/excel/excel.pivotfieldcollection) que tem exatamente um [PivotField](/javascript/api/excel/excel.pivotfield). Se o design se expandir para incluir tabelas dinâmicas OLAP, isso poderá mudar.
- Um [PivotField](/javascript/api/excel/excel.pivotfield) pode ter um ou mais [PivotFilters aplicados](/javascript/api/excel/excel.pivotfilters) , desde que [PivotHierarchy](/javascript/api/excel/excel.pivothierarchy) do campo seja atribuído a uma categoria de hierarquia.
- Um [PivotField](/javascript/api/excel/excel.pivotfield) contém um [PivotItemCollection](/javascript/api/excel/excel.pivotitemcollection) que tem vários [PivotItems](/javascript/api/excel/excel.pivotitem).
- Uma [Tabela Dinâmica](/javascript/api/excel/excel.pivottable) contém um [PivotLayout](/javascript/api/excel/excel.pivotlayout) que define onde os [PivotFields](/javascript/api/excel/excel.pivotfield) e [PivotItems são exibidos](/javascript/api/excel/excel.pivotitem) na planilha. O layout também controla algumas configurações de exibição para a Tabela Dinâmica.

Vejamos como essas relações se aplicam a alguns dados de exemplo. Os dados a seguir descrevem as vendas de frutas de vários farms. Ele será o exemplo ao longo deste artigo.

![Uma coleção de vendas de frutas de diferentes tipos de farms diferentes.](../images/excel-pivots-raw-data.png)

Esses dados de vendas de farm de frutas serão usados para fazer uma tabela dinâmica. Cada coluna, como **Tipos**, é um `PivotHierarchy`. A **hierarquia Tipos** contém o **campo Tipos** . O **campo Tipos** contém os itens **Apple**, **Kiwi**, **Limão**, **Limão** e **Laranja**.

### <a name="hierarchies"></a>Hierarquias

As Tabelas Dinâmicas são organizadas com base em quatro categorias de hierarquia: [linha](/javascript/api/excel/excel.rowcolumnpivothierarchy), [coluna](/javascript/api/excel/excel.rowcolumnpivothierarchy), [dados](/javascript/api/excel/excel.datapivothierarchy) e [filtro](/javascript/api/excel/excel.filterpivothierarchy).

Os dados do farm mostrados anteriormente têm cinco hierarquias: **Farms**, **Type**, **Classification**, **Crates Sold at Farm** e **Crates Sold Wholesale**. Cada hierarquia só pode existir em uma das quatro categorias. Se **Type** for adicionado a hierarquias de coluna, ele também não poderá estar na linha, dados ou hierarquias de filtro. Se **Type** for subsequentemente adicionado às hierarquias de linha, ele será removido das hierarquias de coluna. Esse comportamento é o mesmo se a atribuição de hierarquia é feita por meio da interface do usuário Excel ou do Excel APIs JavaScript.

Hierarquias de linhas e colunas definem como os dados serão agrupados. Por exemplo, uma hierarquia de linhas **de Farms** agrupa todos os conjuntos de dados do mesmo farm. A escolha entre a hierarquia de linha e coluna define a orientação da Tabela Dinâmica.

Hierarquias de dados são os valores a serem agregados com base nas hierarquias de linha e coluna. Uma Tabela Dinâmica com uma hierarquia de linhas de **Farms** e uma hierarquia de dados de **Engradados Vendidos por** Atacado mostra a soma total (por padrão) de todas as diferentes frutas para cada farm.

Hierarquias de filtro incluem ou excluem dados do pivô com base nos valores dentro desse tipo filtrado. Uma hierarquia de filtro de **Classificação com** o tipo **Organic** selecionado mostra apenas dados para frutas orgânicas.

Aqui estão os dados do farm novamente, juntamente com uma tabela dinâmica. A Tabela Dinâmica está usando **Farm** e **Type** como hierarquias de linha, Caixas **Vendidas no Farm** e Engradados Vendidos por Atacado como **hierarquias** de dados (com a função de agregação padrão de soma) e Classificação como  uma hierarquia de filtro (com **a seleção** orgânica).

![Uma seleção de dados de vendas de frutas ao lado de uma tabela dinâmica com hierarquias de linha, dados e filtro.](../images/excel-pivot-table-and-data.png)

Essa Tabela Dinâmica pode ser gerada por meio da API JavaScript ou por meio da interface do usuário Excel usuário. Ambas as opções permitem mais manipulação por meio de complementos.

## <a name="create-a-pivottable"></a>Criar uma tabela dinâmica

As Tabelas Dinâmicas precisam de um nome, fonte e destino. A origem pode ser um endereço de intervalo ou um nome de tabela (passado como `Range`, `string`ou `Table` tipo). O destino é um endereço de intervalo (dado como a `Range` ou `string`).
Os exemplos a seguir mostram várias técnicas de criação de tabela dinâmica.

### <a name="create-a-pivottable-with-range-addresses"></a>Criar uma tabela dinâmica com endereços de intervalo

```js
await Excel.run(async (context) => {
    // Create a PivotTable named "Farm Sales" on the current worksheet at cell
    // A22 with data from the range A1:E21.
    context.workbook.worksheets.getActiveWorksheet().pivotTables.add(
      "Farm Sales", "A1:E21", "A22");

    await context.sync();
});
```

### <a name="create-a-pivottable-with-range-objects"></a>Criar uma tabela dinâmica com objetos Range

```js
await Excel.run(async (context) => {
    // Create a PivotTable named "Farm Sales" on a worksheet called "PivotWorksheet" at cell A2
    // the data comes from the worksheet "DataWorksheet" across the range A1:E21.
    let rangeToAnalyze = context.workbook.worksheets.getItem("DataWorksheet").getRange("A1:E21");
    let rangeToPlacePivot = context.workbook.worksheets.getItem("PivotWorksheet").getRange("A2");
    context.workbook.worksheets.getItem("PivotWorksheet").pivotTables.add(
      "Farm Sales", rangeToAnalyze, rangeToPlacePivot);

    await context.sync();
});
```

### <a name="create-a-pivottable-at-the-workbook-level"></a>Criar uma Tabela Dinâmica no nível da workbook

```js
await Excel.run(async (context) => {
    // Create a PivotTable named "Farm Sales" on a worksheet called "PivotWorksheet" at cell A2
    // the data is from the worksheet "DataWorksheet" across the range A1:E21.
    context.workbook.pivotTables.add(
        "Farm Sales", "DataWorksheet!A1:E21", "PivotWorksheet!A2");

    await context.sync();
});
```

## <a name="use-an-existing-pivottable"></a>Usar uma tabela dinâmica existente

As Tabelas Dinâmicas criadas manualmente também são acessíveis por meio da coleção PivotTable da pasta de trabalho ou de planilhas individuais. O código a seguir obtém uma Tabela Dinâmica chamada **Meu Pivô** da lista de trabalho.

```js
await Excel.run(async (context) => {
    let pivotTable = context.workbook.pivotTables.getItem("My Pivot");
    await context.sync();
});
```

## <a name="add-rows-and-columns-to-a-pivottable"></a>Adicionar linhas e colunas a uma tabela dinâmica

Linhas e colunas giram os dados em torno dos valores desses campos.

A **adição da coluna Farm** gira todas as vendas ao redor de cada farm. Adicionar as **linhas Tipo** e **Classificação** quebra ainda mais os dados com base em quais frutas foram vendidas e se foram orgânicas ou não.

![Uma tabela dinâmica com uma coluna farm e linhas Tipo e Classificação.](../images/excel-pivots-table-rows-and-columns.png)

```js
await Excel.run(async (context) => {
    let pivotTable = context.workbook.worksheets.getActiveWorksheet().pivotTables.getItem("Farm Sales");

    pivotTable.rowHierarchies.add(pivotTable.hierarchies.getItem("Type"));
    pivotTable.rowHierarchies.add(pivotTable.hierarchies.getItem("Classification"));

    pivotTable.columnHierarchies.add(pivotTable.hierarchies.getItem("Farm"));

    await context.sync();
});
```

Você também pode ter uma Tabela Dinâmica com apenas linhas ou colunas.

```js
await Excel.run(async (context) => {
    let pivotTable = context.workbook.worksheets.getActiveWorksheet().pivotTables.getItem("Farm Sales");
    pivotTable.rowHierarchies.add(pivotTable.hierarchies.getItem("Farm"));
    pivotTable.rowHierarchies.add(pivotTable.hierarchies.getItem("Type"));
    pivotTable.rowHierarchies.add(pivotTable.hierarchies.getItem("Classification"));

    await context.sync();
});
```

## <a name="add-data-hierarchies-to-the-pivottable"></a>Adicionar hierarquias de dados à Tabela Dinâmica

Hierarquias de dados preenchem a Tabela Dinâmica com informações para combinar com base nas linhas e colunas. A adição das hierarquias de dados de **Caixas Vendidas no Farm** e caixas **vendidas por** atacado fornece somas desses números para cada linha e coluna.

No exemplo **, Farm e** **Type** são linhas, com as vendas do engradado como os dados.

![Uma Tabela Dinâmica mostrando o total de vendas de diferentes frutas com base no farm de onde eles vieram.](../images/excel-pivots-data-hierarchy.png)

```js
await Excel.run(async (context) => {
    let pivotTable = context.workbook.worksheets.getActiveWorksheet().pivotTables.getItem("Farm Sales");

    // "Farm" and "Type" are the hierarchies on which the aggregation is based.
    pivotTable.rowHierarchies.add(pivotTable.hierarchies.getItem("Farm"));
    pivotTable.rowHierarchies.add(pivotTable.hierarchies.getItem("Type"));

    // "Crates Sold at Farm" and "Crates Sold Wholesale" are the hierarchies
    // that will have their data aggregated (summed in this case).
    pivotTable.dataHierarchies.add(pivotTable.hierarchies.getItem("Crates Sold at Farm"));
    pivotTable.dataHierarchies.add(pivotTable.hierarchies.getItem("Crates Sold Wholesale"));

    await context.sync();
});
```

## <a name="pivottable-layouts-and-getting-pivoted-data"></a>Layouts de tabela dinâmica e informações dinâmicas

Um [PivotLayout](/javascript/api/excel/excel.pivotlayout) define o posicionamento das hierarquias e seus dados. Você acessa o layout para determinar os intervalos onde os dados são armazenados.

O diagrama a seguir mostra quais chamadas de função de layout correspondem a quais intervalos da Tabela Dinâmica.

![Um diagrama mostrando quais seções de uma tabela dinâmica são retornadas pelas funções de intervalo de obter do layout.](../images/excel-pivots-layout-breakdown.png)

### <a name="get-data-from-the-pivottable"></a>Obter dados da tabela dinâmica

O layout define como a Tabela Dinâmica é exibida na planilha. Isso significa que o `PivotLayout` objeto controla os intervalos usados para elementos de tabela dinâmica. Use os intervalos fornecidos pelo layout para obter dados coletados e agregados pela Tabela Dinâmica. Em particular, use `PivotLayout.getDataBodyRange` para acessar os dados produzidos pela Tabela Dinâmica.

O código a seguir demonstra como obter a última linha dos dados de tabela dinâmica passando pelo layout (o **Grande Total** da Soma de Caixas **Vendidas no Farm** e a Soma das **colunas De Engradados Vendidos** por Atacado no exemplo anterior). Esses valores são, em seguida, resumidos para um total final, que é exibido na célula **E30** (fora da Tabela Dinâmica).

```js
await Excel.run(async (context) => {
    let pivotTable = context.workbook.worksheets.getActiveWorksheet().pivotTables.getItem("Farm Sales");

    // Get the totals for each data hierarchy from the layout.
    let range = pivotTable.layout.getDataBodyRange();
    let grandTotalRange = range.getLastRow();
    grandTotalRange.load("address");
    await context.sync();

    // Sum the totals from the PivotTable data hierarchies and place them in a new range, outside of the PivotTable.
    let masterTotalRange = context.workbook.worksheets.getActiveWorksheet().getRange("E30");
    masterTotalRange.formulas = [["=SUM(" + grandTotalRange.address + ")"]];
    await context.sync();
});
```

### <a name="layout-types"></a>Tipos de layout

As Tabelas Dinâmicas têm três estilos de layout: Compact, Outline e Tabular. Vimos o estilo compacto nos exemplos anteriores.

Os exemplos a seguir usam os estilos de contorno e tabular, respectivamente. O exemplo de código mostra como fazer o ciclo entre os diferentes layouts.

#### <a name="outline-layout"></a>Layout de estrutura de estrutura

![Uma tabela dinâmica usando o layout de estrutura de estrutura.](../images/excel-pivots-outline-layout.png)

#### <a name="tabular-layout"></a>Layout tabular

![Uma tabela dinâmica usando o layout tabular.](../images/excel-pivots-tabular-layout.png)

#### <a name="pivotlayout-type-switch-code-sample"></a>Exemplo de código de opção do tipo PivotLayout

```js
await Excel.run(async (context) => {
    // Change the PivotLayout.type to a new type.
    let pivotTable = context.workbook.worksheets.getActiveWorksheet().pivotTables.getItem("Farm Sales");
    pivotTable.layout.load("layoutType");
    await context.sync();

    // Cycle between the three layout types.
    if (pivotTable.layout.layoutType === "Compact") {
        pivotTable.layout.layoutType = "Outline";
    } else if (pivotTable.layout.layoutType === "Outline") {
        pivotTable.layout.layoutType = "Tabular";
    } else {
        pivotTable.layout.layoutType = "Compact";
    }

    await context.sync();
});
```

### <a name="other-pivotlayout-functions"></a>Outras funções PivotLayout

Por padrão, tabelas dinâmicas ajustam tamanhos de linha e coluna conforme necessário. Isso é feito quando a Tabela Dinâmica é atualizada. `PivotLayout.autoFormat` especifica esse comportamento. Qualquer alteração de tamanho de linha ou coluna feita pelo seu complemento persiste quando `autoFormat` é `false`. Além disso, as configurações padrão de uma tabela dinâmica mantêm qualquer formatação personalizada na Tabela Dinâmica (como preenchimentos e alterações de fonte). Definir `PivotLayout.preserveFormatting` para `false` aplicar o formato padrão quando atualizado.

Um `PivotLayout` também controla as configurações de header e de linha total, como as células de dados vazias são exibidas e [as opções de texto alt](https://support.microsoft.com/topic/44989b2a-903c-4d9a-b742-6a75b451c669) . A [referência PivotLayout](/javascript/api/excel/excel.pivotlayout) fornece uma lista completa desses recursos.

O exemplo de código `"--"`a seguir faz com que as células de dados vazias exibem a cadeia de caracteres , formate o intervalo do corpo para um alinhamento horizontal consistente e garante que as alterações de formatação permaneçam mesmo após a atualização da Tabela Dinâmica.

```js
await Excel.run(async (context) => {
    let pivotTable = context.workbook.pivotTables.getItem("Farm Sales");
    let pivotLayout = pivotTable.layout;

    // Set a default value for an empty cell in the PivotTable. This doesn't include cells left blank by the layout.
    pivotLayout.emptyCellText = "--";

    // Set the text alignment to match the rest of the PivotTable.
    pivotLayout.getDataBodyRange().format.horizontalAlignment = Excel.HorizontalAlignment.right;

    // Ensure empty cells are filled with a default value.
    pivotLayout.fillEmptyCells = true;

    // Ensure that the format settings persist, even after the PivotTable is refreshed and recalculated.
    pivotLayout.preserveFormatting = true;
    await context.sync();
});
```

## <a name="delete-a-pivottable"></a>Excluir uma tabela dinâmica

Tabelas Dinâmicas são excluídas usando seu nome.

```js
await Excel.run(async (context) => {
    context.workbook.worksheets.getItem("Pivot").pivotTables.getItem("Farm Sales").delete();
    await context.sync();
});
```

## <a name="filter-a-pivottable"></a>Filtrar uma tabela dinâmica

O método principal para filtrar dados de tabela dinâmica é com PivotFilters. As slicers oferecem um método alternativo de filtragem menos flexível.

[PivotFilters](/javascript/api/excel/excel.pivotfilters) filtram dados com base nas quatro categorias de hierarquia de uma tabela [dinâmica (](#hierarchies) filtros, colunas, linhas e valores). Há quatro tipos de PivotFilters, permitindo filtragem baseada em data de calendário, análise de cadeia de caracteres, comparação de números e filtragem com base em uma entrada personalizada.

[As slicers](/javascript/api/excel/excel.slicer) podem ser aplicadas a tabelas dinâmicas e Excel regulares. Quando aplicada a uma Tabela Dinâmica, as slicers funcionam como um [PivotManualFilter](#pivotmanualfilter) e permitem a filtragem com base em uma entrada personalizada. Ao contrário dos PivotFilters, as slicers têm um [Excel de interface do usuário](https://support.microsoft.com/office/249f966b-a9d5-4b0f-b31a-12651785d29d). Com a `Slicer` classe, você cria esse componente de interface do usuário, gerencia a filtragem e controla sua aparência visual.

### <a name="filter-with-pivotfilters"></a>Filtrar com PivotFilters

[Os PivotFilters](/javascript/api/excel/excel.pivotfilters) permitem filtrar dados de tabela dinâmica com base nas quatro categorias de [hierarquia (](#hierarchies) filtros, colunas, linhas e valores). No modelo de objeto pivotTable, `PivotFilters` são aplicados a um [PivotField](/javascript/api/excel/excel.pivotfield) e cada `PivotField` um pode ter um ou mais atribuídos `PivotFilters`. Para aplicar PivotFilters a um PivotField, a [PivotHierarchy](/javascript/api/excel/excel.pivothierarchy) correspondente do campo deve ser atribuída a uma categoria de hierarquia.

#### <a name="types-of-pivotfilters"></a>Tipos de PivotFilters

| Tipo de filtro | Finalidade de filtro | Referência da API JavaScript do Excel |
|:--- |:--- |:--- |
| DateFilter | Filtragem baseada em data de calendário. | [PivotDateFilter](/javascript/api/excel/excel.pivotdatefilter) |
| LabelFilter | Filtragem de comparação de texto. | [PivotLabelFilter](/javascript/api/excel/excel.pivotlabelfilter) |
| ManualFilter | Filtragem de entrada personalizada. | [PivotManualFilter](/javascript/api/excel/excel.pivotmanualfilter) |
| ValueFilter | Filtragem de comparação de números. | [PivotValueFilter](/javascript/api/excel/excel.pivotvaluefilter) |

#### <a name="create-a-pivotfilter"></a>Criar um PivotFilter

Para filtrar dados de tabela dinâmica com um `Pivot*Filter` (como um `PivotDateFilter`), aplique o filtro a um [PivotField](/javascript/api/excel/excel.pivotfield). Os quatro exemplos de código a seguir mostram como usar cada um dos quatro tipos de PivotFilters.

##### <a name="pivotdatefilter"></a>PivotDateFilter

O primeiro exemplo de código aplica um [PivotDateFilter](/javascript/api/excel/excel.pivotdatefilter) ao PivotField Atualizado de Data, ocultando todos os dados antes **de 2020-08-01**.

> [!IMPORTANT]
> Um `Pivot*Filter` não pode ser aplicado a um PivotField, a menos que PivotHierarchy desse campo seja atribuída a uma categoria de hierarquia. No exemplo de código a seguir, o `dateHierarchy` deve ser adicionado `rowHierarchies` à categoria da tabela dinâmica antes de poder ser usado para filtragem.

```js
await Excel.run(async (context) => {
    // Get the PivotTable and the date hierarchy.
    let pivotTable = context.workbook.worksheets.getActiveWorksheet().pivotTables.getItem("Farm Sales");
    let dateHierarchy = pivotTable.rowHierarchies.getItemOrNullObject("Date Updated");
    await context.sync();

    // PivotFilters can only be applied to PivotHierarchies that are being used for pivoting.
    // If it's not already there, add "Date Updated" to the hierarchies.
    if (dateHierarchy.isNullObject) {
        dateHierarchy = pivotTable.rowHierarchies.add(pivotTable.hierarchies.getItem("Date Updated"));
    }

    // Apply a date filter to filter out anything logged before August.
    let filterField = dateHierarchy.fields.getItem("Date Updated");
    let dateFilter = {
        condition: Excel.DateFilterCondition.afterOrEqualTo,
        comparator: {
        date: "2020-08-01",
        specificity: Excel.FilterDatetimeSpecificity.month
        }
    };
    filterField.applyFilter({ dateFilter: dateFilter });
    
    await context.sync();
});
```

> [!NOTE]
> Os três trechos de código a seguir exibem apenas trechos específicos do filtro, em vez de chamadas completas `Excel.run` .

##### <a name="pivotlabelfilter"></a>PivotLabelFilter

O segundo trecho de código demonstra como aplicar um [PivotLabelFilter](/javascript/api/excel/excel.pivotlabelfilter) ao **Tipo** PivotField, `LabelFilterCondition.beginsWith` usando a propriedade para excluir rótulos que começam com a letra **L**.

```js
    // Get the "Type" field.
    let filterField = pivotTable.hierarchies.getItem("Type").fields.getItem("Type");

    // Filter out any types that start with "L" ("Lemons" and "Limes" in this case).
    let filter: Excel.PivotLabelFilter = {
      condition: Excel.LabelFilterCondition.beginsWith,
      substring: "L",
      exclusive: true
    };

    // Apply the label filter to the field.
    filterField.applyFilter({ labelFilter: filter });
```

##### <a name="pivotmanualfilter"></a>PivotManualFilter

O terceiro trecho de código aplica um filtro manual com [PivotManualFilter](/javascript/api/excel/excel.pivotmanualfilter) ao campo Classificação, filtrando dados que não incluem a classificação **Orgânica**.

```js
    // Apply a manual filter to include only a specific PivotItem (the string "Organic").
    let filterField = classHierarchy.fields.getItem("Classification");
    let manualFilter = { selectedItems: ["Organic"] };
    filterField.applyFilter({ manualFilter: manualFilter });
```

##### <a name="pivotvaluefilter"></a>PivotValueFilter

Para comparar números, use um filtro de valor com [PivotValueFilter](/javascript/api/excel/excel.pivotvaluefilter), conforme mostrado no trecho de código final. O `PivotValueFilter` compara os dados no **Farm PivotField** com os dados no **Campo** Pivô de Venda de Engradados, incluindo apenas farms cuja soma de caixas vendidas excede o valor **500**.

```js
    // Get the "Farm" field.
    let filterField = pivotTable.hierarchies.getItem("Farm").fields.getItem("Farm");
    
    // Filter to only include rows with more than 500 wholesale crates sold.
    let filter: Excel.PivotValueFilter = {
      condition: Excel.ValueFilterCondition.greaterThan,
      comparator: 500,
      value: "Sum of Crates Sold Wholesale"
    };
    
    // Apply the value filter to the field.
    filterField.applyFilter({ valueFilter: filter });
```

#### <a name="remove-pivotfilters"></a>Remover PivotFilters

Para remover todos os PivotFilters, aplique o `clearAllFilters` método a cada PivotField, conforme mostrado no exemplo de código a seguir.

```js
await Excel.run(async (context) => {
    // Get the PivotTable.
    let pivotTable = context.workbook.worksheets.getActiveWorksheet().pivotTables.getItem("Farm Sales");
    pivotTable.hierarchies.load("name");
    await context.sync();

    // Clear the filters on each PivotField.
    pivotTable.hierarchies.items.forEach(function (hierarchy) {
        hierarchy.fields.getItem(hierarchy.name).clearAllFilters();
    });
    await context.sync();
});
```

### <a name="filter-with-slicers"></a>Filtrar com slicers

[As slicers](/javascript/api/excel/excel.slicer) permitem que os dados sejam filtrados de uma tabela Excel dinâmica. Uma slicer usa valores de uma coluna especificada ou PivotField para filtrar linhas correspondentes. Esses valores são armazenados [como objetos SlicerItem](/javascript/api/excel/excel.sliceritem) no `Slicer`. Seu complemento pode ajustar esses filtros, assim como os usuários (por meio [da interface do usuário Excel interface do usuário](https://support.microsoft.com/office/249f966b-a9d5-4b0f-b31a-12651785d29d)). A slicer fica na parte superior da planilha na camada de desenho, conforme mostrado na captura de tela a seguir.

![Uma filtragem de dados de uma slicer em uma tabela dinâmica.](../images/excel-slicer.png)

> [!NOTE]
> As técnicas descritas nesta seção se concentram em como usar slicers conectados a Tabelas Dinâmicas. As mesmas técnicas também se aplicam ao uso de slicers conectados a tabelas.

#### <a name="create-a-slicer"></a>Criar uma slicer

Você pode criar uma slicer em uma pasta de trabalho ou planilha usando o `Workbook.slicers.add` método ou o `Worksheet.slicers.add` método. Isso adiciona uma slicer à [SlicerCollection](/javascript/api/excel/excel.slicercollection) do objeto `Workbook` especificado ou `Worksheet` . O `SlicerCollection.add` método tem três parâmetros:

- `slicerSource`: A fonte de dados na qual a nova slicer é baseada. Pode ser uma `PivotTable`cadeia de caracteres , `Table`ou que representa o nome ou a ID de um `PivotTable` ou `Table`.
- `sourceField`: O campo na fonte de dados pelo qual filtrar. Pode ser uma `PivotField`cadeia de caracteres , `TableColumn`ou que representa o nome ou a ID de um `PivotField` ou `TableColumn`.
- `slicerDestination`: A planilha onde a nova slicer será criada. Pode ser um objeto `Worksheet` ou o nome ou a ID de um `Worksheet`. Esse parâmetro é desnecessário quando o `SlicerCollection` é acessado por meio de `Worksheet.slicers`. Nesse caso, a planilha da coleção é usada como destino.

O exemplo de código a seguir adiciona uma nova slicer à **planilha** Dinâmica. A origem da slicer é a **Tabela** Dinâmica de Vendas do Farm e filtra usando os **dados Type** . A slicer também é chamada **de Fruit Slicer** para referência futura.

```js
await Excel.run(async (context) => {
    let sheet = context.workbook.worksheets.getItem("Pivot");
    let slicer = sheet.slicers.add(
        "Farm Sales" /* The slicer data source. For PivotTables, this can be the PivotTable object reference or name. */,
        "Type" /* The field in the data to filter by. For PivotTables, this can be a PivotField object reference or ID. */
    );
    slicer.name = "Fruit Slicer";
    await context.sync();
});
```

#### <a name="filter-items-with-a-slicer"></a>Filtrar itens com uma slicer

A slicer filtra a Tabela Dinâmica com itens do `sourceField`. O `Slicer.selectItems` método define os itens que permanecem na slicer. Esses itens são passados para o método como um `string[]`, representando as chaves dos itens. Todas as linhas que contêm esses itens permanecem na agregação da tabela dinâmica. Chamadas subsequentes para `selectItems` definir a lista como as chaves especificadas nessas chamadas.

> [!NOTE]
> Se `Slicer.selectItems` for passado um item que não está na fonte de dados, será `InvalidArgument` lançado um erro. O conteúdo pode ser verificado por meio da `Slicer.slicerItems` propriedade, que é [uma SlicerItemCollection](/javascript/api/excel/excel.sliceritemcollection).

O exemplo de código a seguir mostra três itens que estão sendo selecionados para a slicer: **Limão**, **Limão** e **Laranja**.

```js
await Excel.run(async (context) => {
    let slicer = context.workbook.slicers.getItem("Fruit Slicer");
    // Anything other than the following three values will be filtered out of the PivotTable for display and aggregation.
    slicer.selectItems(["Lemon", "Lime", "Orange"]);
    await context.sync();
});
```

Para remover todos os filtros da slicer, use o `Slicer.clearFilters` método, conforme mostrado no exemplo a seguir.

```js
await Excel.run(async (context) => {
    let slicer = context.workbook.slicers.getItem("Fruit Slicer");
    slicer.clearFilters();
    await context.sync();
});
```

#### <a name="style-and-format-a-slicer"></a>Estilo e formatar uma slicer

Você pode ajustar as configurações de exibição de uma slicer por meio de `Slicer` propriedades. O exemplo de código a seguir define o estilo como **SlicerStyleLight6**, define o texto na parte superior da slicer como Tipos de **Frutas, coloca** a slicer na posição **(395, 15)** na camada de desenho e define o tamanho da slicer como **135x150** pixels.

```js
await Excel.run(async (context) => {
    let slicer = context.workbook.slicers.getItem("Fruit Slicer");
    slicer.caption = "Fruit Types";
    slicer.left = 395;
    slicer.top = 15;
    slicer.height = 135;
    slicer.width = 150;
    slicer.style = "SlicerStyleLight6";
    await context.sync();
});
```

#### <a name="delete-a-slicer"></a>Excluir uma slicer

Para excluir uma slicer, chame o `Slicer.delete` método. O exemplo de código a seguir exclui a primeira fatia da planilha atual.

```js
await Excel.run(async (context) => {
    let sheet = context.workbook.worksheets.getActiveWorksheet();
    sheet.slicers.getItemAt(0).delete();
    await context.sync();
});
```

## <a name="change-aggregation-function"></a>Função Alterar agregação

Hierarquias de dados têm seus valores agregados. Para conjuntos de dados de números, essa é uma soma por padrão. A `summarizeBy` propriedade define esse comportamento com base em um [tipo AggregationFunction](/javascript/api/excel/excel.aggregationfunction) .

Os tipos de função de agregação `Sum`atualmente suportados são , `Count`, `Average`, , `Max`, `Min`, `Product`, `CountNumbers`, , `StandardDeviation`, `StandardDeviationP`, , `Variance`, , `VarianceP`e `Automatic` (o padrão).

Os exemplos de código a seguir modificam a agregação como médias dos dados.

```js
await Excel.run(async (context) => {
    let pivotTable = context.workbook.worksheets.getActiveWorksheet().pivotTables.getItem("Farm Sales");
    pivotTable.dataHierarchies.load("no-properties-needed");
    await context.sync();

    // Change the aggregation from the default sum to an average of all the values in the hierarchy.
    pivotTable.dataHierarchies.items[0].summarizeBy = Excel.AggregationFunction.average;
    pivotTable.dataHierarchies.items[1].summarizeBy = Excel.AggregationFunction.average;
    await context.sync();
});
```

## <a name="change-calculations-with-a-showasrule"></a>Alterar cálculos com um ShowAsRule

As Tabelas Dinâmicas, por padrão, agregam os dados de suas hierarquias de linha e coluna independentemente. Um [ShowAsRule](/javascript/api/excel/excel.showasrule) altera a hierarquia de dados para valores de saída com base em outros itens na Tabela Dinâmica.

O `ShowAsRule` objeto tem três propriedades:

- `calculation`: O tipo de cálculo relativo a ser aplicado à hierarquia de dados (o padrão é `none`).
- `baseField`: [PivotField](/javascript/api/excel/excel.pivotfield) na hierarquia que contém os dados base antes da aplicação do cálculo. Como Excel tabelas dinâmicas têm um mapeamento de hierarquia para campo, você usará o mesmo nome para acessar a hierarquia e o campo.
- `baseItem`: [PivotItem](/javascript/api/excel/excel.pivotitem) individual comparado com os valores dos campos base com base no tipo de cálculo. Nem todos os cálculos exigem esse campo.

O exemplo a seguir define o cálculo na **Soma de Caixas Vendidas na hierarquia** de dados do Farm como uma porcentagem do total da coluna.
Ainda queremos que a granularidade se estenda até o nível de tipo de frutas, portanto, vamos usar a hierarquia **de linha Type** e seu campo subjacente.
O exemplo também tem **Farm** como a hierarquia da primeira linha, portanto, o total de entradas do farm exibe a porcentagem que cada farm também é responsável por produzir.

![Uma Tabela Dinâmica mostrando as porcentagens de vendas de frutas em relação ao total geral para farms individuais e tipos de frutas individuais em cada farm.](../images/excel-pivots-showas-percentage.png)

```js
await Excel.run(async (context) => {
    let pivotTable = context.workbook.worksheets.getActiveWorksheet().pivotTables.getItem("Farm Sales");
    let farmDataHierarchy = pivotTable.dataHierarchies.getItem("Sum of Crates Sold at Farm");

    farmDataHierarchy.load("showAs");
    await context.sync();

    // Show the crates of each fruit type sold at the farm as a percentage of the column's total.
    let farmShowAs = farmDataHierarchy.showAs;
    farmShowAs.calculation = Excel.ShowAsCalculation.percentOfColumnTotal;
    farmShowAs.baseField = pivotTable.rowHierarchies.getItem("Type").fields.getItem("Type");
    farmDataHierarchy.showAs = farmShowAs;
    farmDataHierarchy.name = "Percentage of Total Farm Sales";
});
```

O exemplo anterior definiu o cálculo como a coluna, em relação ao campo de uma hierarquia de linha individual. Quando o cálculo se refere a um item individual, use a `baseItem` propriedade.

O exemplo a seguir mostra o `differenceFrom` cálculo. Ele exibe a diferença das entradas de hierarquia de dados de vendas de caixa de farm em relação às de **Farms.**
O `baseField` é **Farm**, portanto, vemos as diferenças entre os outros farms, bem como as repartições para cada tipo de frutas como (**Type** também é uma hierarquia de linhas neste exemplo).

![Uma Tabela Dinâmica mostrando as diferenças de vendas de frutas entre "A Farms" e as outras. Isso mostra a diferença no total de vendas de frutas dos farms e nas vendas de tipos de frutas. Se "A Farms" não vender um tipo específico de frutas, "#N/A" será exibido.](../images/excel-pivots-showas-differencefrom.png)

```js
await Excel.run(async (context) => {
    let pivotTable = context.workbook.worksheets.getActiveWorksheet().pivotTables.getItem("Farm Sales");
    let farmDataHierarchy = pivotTable.dataHierarchies.getItem("Sum of Crates Sold at Farm");

    farmDataHierarchy.load("showAs");
    await context.sync();
        
    // Show the difference between crate sales of the "A Farms" and the other farms.
    // This difference is both aggregated and shown for individual fruit types (where applicable).
    let farmShowAs = farmDataHierarchy.showAs;
    farmShowAs.calculation = Excel.ShowAsCalculation.differenceFrom;
    farmShowAs.baseField = pivotTable.rowHierarchies.getItem("Farm").fields.getItem("Farm");
    farmShowAs.baseItem = pivotTable.rowHierarchies.getItem("Farm").fields.getItem("Farm").items.getItem("A Farms");
    farmDataHierarchy.showAs = farmShowAs;
    farmDataHierarchy.name = "Difference from A Farms";
});
```

## <a name="change-hierarchy-names"></a>Alterar nomes de hierarquia

Os campos de hierarquia são editáveis. O código a seguir demonstra como alterar os nomes exibidos de duas hierarquias de dados.

```js
await Excel.run(async (context) => {
    let dataHierarchies = context.workbook.worksheets.getActiveWorksheet()
        .pivotTables.getItem("Farm Sales").dataHierarchies;
    dataHierarchies.load("no-properties-needed");
    await context.sync();

    // Changing the displayed names of these entries.
    dataHierarchies.items[0].name = "Farm Sales";
    dataHierarchies.items[1].name = "Wholesale";
});
```

## <a name="see-also"></a>Confira também

- [Modelo de objeto JavaScript do Excel em Suplementos do Office](excel-add-ins-core-concepts.md)
- [Excel referência da API JavaScript](/javascript/api/excel)
