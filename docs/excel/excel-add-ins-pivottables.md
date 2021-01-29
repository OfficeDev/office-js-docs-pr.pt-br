---
title: Trabalhar com Tabelas Dinâmicas usando a API JavaScript do Excel
description: Use a API JavaScript do Excel para criar Tabelas Dinâmicas e interagir com seus componentes.
ms.date: 01/26/2021
localization_priority: Normal
ms.openlocfilehash: 9832322d40bbeb247685ff2498bdce42975c0377
ms.sourcegitcommit: 3123b9819c5225ee45a5312f64be79e46cbd0e3c
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 01/29/2021
ms.locfileid: "50043908"
---
# <a name="work-with-pivottables-using-the-excel-javascript-api"></a>Trabalhar com Tabelas Dinâmicas usando a API JavaScript do Excel

As Tabelas Dinâmicas simplificam conjuntos de dados maiores. Eles permitem a manipulação rápida de dados agrupados. A API JavaScript do Excel permite que seu complemento crie Tabelas Dinâmicas e interaja com seus componentes. Este artigo descreve como as Tabelas Dinâmicas são representadas pela API JavaScript do Office e fornece exemplos de código para cenários principais.

Se você não estiver familiarizado com a funcionalidade de Tabelas Dinâmicas, considere explorá-las como um usuário final.
Consulte [Criar uma Tabela Dinâmica para analisar dados de planilha](https://support.office.com/article/Import-and-analyze-data-ccd3c4a6-272f-4c97-afbb-d3f27407fcde#ID0EAABAAA=PivotTables) para ter uma boa cartilha nessas ferramentas.

> [!IMPORTANT]
> As Tabelas Dinâmicas criadas com OLAP não têm suporte no momento. Também não há suporte para o Power Pivot.

## <a name="object-model"></a>Modelo de objetos

A [Tabela Dinâmica](/javascript/api/excel/excel.pivottable) é o objeto central para Tabelas Dinâmicas na API JavaScript do Office.

- `Workbook.pivotTables` e `Worksheet.pivotTables` são [PivotTableCollections](/javascript/api/excel/excel.pivottablecollection) que contêm as [Tabelas Dinâmicas](/javascript/api/excel/excel.pivottable) na pasta de trabalho e na planilha, respectivamente.
- Uma [tabela dinâmica](/javascript/api/excel/excel.pivottable) contém um [PivotHierarchyCollection](/javascript/api/excel/excel.pivothierarchycollection) que tem várias [PivotHierarchies](/javascript/api/excel/excel.pivothierarchy).
- Essas [PivotHierarchies](/javascript/api/excel/excel.pivothierarchy) podem ser adicionadas a coleções de hierarquias específicas para definir como os dados de tabela dinâmica (conforme explicado na [seção a seguir).](#hierarchies)
- Um [PivotHierarchy](/javascript/api/excel/excel.pivothierarchy) contém [um PivotFieldCollection](/javascript/api/excel/excel.pivotfieldcollection) que tem exatamente um [PivotField](/javascript/api/excel/excel.pivotfield). Se o design se expandir para incluir tabelas dinâmicas OLAP, isso poderá mudar.
- Um [PivotField](/javascript/api/excel/excel.pivotfield) pode ter um ou mais [PivotFilters aplicados,](/javascript/api/excel/excel.pivotfilters) desde que [PivotHierarchy](/javascript/api/excel/excel.pivothierarchy) do campo seja atribuído a uma categoria de hierarquia. 
- Um [PivotField](/javascript/api/excel/excel.pivotfield) contém [um PivotItemCollection](/javascript/api/excel/excel.pivotitemcollection) que tem vários [PivotItems](/javascript/api/excel/excel.pivotitem).
- Uma [Tabela Dinâmica](/javascript/api/excel/excel.pivottable) contém um [PivotLayout](/javascript/api/excel/excel.pivotlayout) que define onde [pivotFields](/javascript/api/excel/excel.pivotfield) e [PivotItems](/javascript/api/excel/excel.pivotitem) são exibidos na planilha.

Vamos ver como essas relações se aplicam a alguns dados de exemplo. Os dados a seguir descrevem as vendas de árvores de vários farms. Ele será o exemplo ao longo deste artigo.

![Uma coleção de vendas de árvores de diferentes tipos de farms diferentes.](../images/excel-pivots-raw-data.png)

Esses dados de vendas de farm de árvores serão usados para fazer uma Tabela Dinâmica. Cada coluna, como **Tipos,** é um `PivotHierarchy` . A **hierarquia Tipos** contém o campo **Tipos.** O **campo Tipos** contém os itens **Apple**, **Kiwi**, **Limão**, **Verde-limão** e **Laranja**.

### <a name="hierarchies"></a>Hierarquias

As Tabelas Dinâmicas são organizadas com base em quatro categorias de hierarquia: [linha](/javascript/api/excel/excel.rowcolumnpivothierarchy), [coluna,](/javascript/api/excel/excel.rowcolumnpivothierarchy) [dados](/javascript/api/excel/excel.datapivothierarchy)e [filtro.](/javascript/api/excel/excel.filterpivothierarchy)

Os dados do farm mostrados anteriormente têm cinco hierarquias: **Farms**, **Tipo** **,** Classificação , Caixas **vendidas** no farm e **Engradados vendidos em um farm.** Cada hierarquia só pode existir em uma das quatro categorias. Se **Type** for adicionado às hierarquias de colunas, ele também não poderá estar nas hierarquias de linha, dados ou filtro. Se **Type** for adicionado subsequentemente às hierarquias de linhas, ele será removido das hierarquias de colunas. Esse comportamento é o mesmo se a atribuição de hierarquia é feita por meio da interface do usuário do Excel ou das APIs JavaScript do Excel.

Hierarquias de linha e coluna definem como os dados serão agrupados. Por exemplo, uma hierarquia de linhas **de Farms** agrupa todos os conjuntos de dados do mesmo farm. A escolha entre a hierarquia de linha e coluna define a orientação da tabela dinâmica.

Hierarquias de dados são os valores a serem agregados com base nas hierarquias de linha e coluna. Uma Tabela Dinâmica com uma hierarquia de linhas de **Farms** e uma hierarquia de dados de **Engradado** vendido mostra a soma total (por padrão) de todas as diferentes árvores para cada farm.

Hierarquias de filtro incluem ou excluem dados da tabela dinâmica com base em valores dentro desse tipo filtrado. Uma hierarquia de filtros **de Classificação** com o tipo **Organic** selected mostra apenas os dados de árvores orgânicas.

Aqui estão os dados do farm novamente, juntamente com uma Tabela Dinâmica. A Tabela Dinâmica está  usando **Farm** e Tipo como **hierarquias** de linha, Engradados vendidos em farm e engradados **vendidos como hierarquias** de  dados (com a função de agregação padrão de soma) e Classificação como uma hierarquia de filtros (com Produtos Orgânicas selecionados). 

![Uma seleção de dados de vendas de resultado ao lado de uma Tabela Dinâmica com hierarquias de linha, dados e filtro.](../images/excel-pivot-table-and-data.png)

Essa Tabela Dinâmica pode ser gerada por meio da API JavaScript ou da interface do usuário do Excel. Ambas as opções permitem mais manipulação por meio de complementos.

## <a name="create-a-pivottable"></a>Criar uma Tabela Dinâmica

As Tabelas Dinâmicas precisam de um nome, fonte e destino. A fonte pode ser um endereço de intervalo ou um nome de tabela (passado como `Range` um `string` , ou `Table` tipo). O destino é um endereço de intervalo (dado como um `Range` ou `string` ).
Os exemplos a seguir mostram várias técnicas de criação de tabela dinâmica.

### <a name="create-a-pivottable-with-range-addresses"></a>Criar uma tabela dinâmica com endereços de intervalo

```js
Excel.run(function (context) {
    // Create a PivotTable named "Farm Sales" on the current worksheet at cell
    // A22 with data from the range A1:E21.
    context.workbook.worksheets.getActiveWorksheet().pivotTables.add(
      "Farm Sales", "A1:E21", "A22");

    return context.sync();
});
```

### <a name="create-a-pivottable-with-range-objects"></a>Criar uma tabela dinâmica com objetos Range

```js
Excel.run(function (context) {
    // Create a PivotTable named "Farm Sales" on a worksheet called "PivotWorksheet" at cell A2
    // the data comes from the worksheet "DataWorksheet" across the range A1:E21.
    var rangeToAnalyze = context.workbook.worksheets.getItem("DataWorksheet").getRange("A1:E21");
    var rangeToPlacePivot = context.workbook.worksheets.getItem("PivotWorksheet").getRange("A2");
    context.workbook.worksheets.getItem("PivotWorksheet").pivotTables.add(
      "Farm Sales", rangeToAnalyze, rangeToPlacePivot);

    return context.sync();
});
```

### <a name="create-a-pivottable-at-the-workbook-level"></a>Criar uma Tabela Dinâmica no nível da agenda

```js
Excel.run(function (context) {
    // Create a PivotTable named "Farm Sales" on a worksheet called "PivotWorksheet" at cell A2
    // the data is from the worksheet "DataWorksheet" across the range A1:E21.
    context.workbook.pivotTables.add(
        "Farm Sales", "DataWorksheet!A1:E21", "PivotWorksheet!A2");

    return context.sync();
});
```

## <a name="use-an-existing-pivottable"></a>Usar uma tabela dinâmica existente

As Tabelas Dinâmicas criadas manualmente também podem ser acessadas por meio da coleção de Tabelas Dinâmicas da pasta de trabalho ou de planilhas individuais. O código a seguir obtém uma tabela dinâmica **chamada My Pivot** da lista de trabalho.

```js
Excel.run(function (context) {
    var pivotTable = context.workbook.pivotTables.getItem("My Pivot");
    return context.sync();
});
```

## <a name="add-rows-and-columns-to-a-pivottable"></a>Adicionar linhas e colunas a uma Tabela Dinâmica

Linhas e colunas giram os dados em torno dos valores desses campos.

A adição **da coluna Farm** gira todas as vendas em torno de cada farm. Adicionar as **linhas Tipo** **e** Classificação divide ainda mais os dados com base em quais árvores foram vendidas e se foram orgânicas ou não.

![Uma Tabela Dinâmica com uma coluna de Farm e linhas Tipo e Classificação.](../images/excel-pivots-table-rows-and-columns.png)

```js
Excel.run(function (context) {
    var pivotTable = context.workbook.worksheets.getActiveWorksheet().pivotTables.getItem("Farm Sales");

    pivotTable.rowHierarchies.add(pivotTable.hierarchies.getItem("Type"));
    pivotTable.rowHierarchies.add(pivotTable.hierarchies.getItem("Classification"));

    pivotTable.columnHierarchies.add(pivotTable.hierarchies.getItem("Farm"));

    return context.sync();
});
```

Você também pode ter uma Tabela Dinâmica com apenas linhas ou colunas.

```js
Excel.run(function (context) {
    var pivotTable = context.workbook.worksheets.getActiveWorksheet().pivotTables.getItem("Farm Sales");
    pivotTable.rowHierarchies.add(pivotTable.hierarchies.getItem("Farm"));
    pivotTable.rowHierarchies.add(pivotTable.hierarchies.getItem("Type"));
    pivotTable.rowHierarchies.add(pivotTable.hierarchies.getItem("Classification"));

    return context.sync();
});
```

## <a name="add-data-hierarchies-to-the-pivottable"></a>Adicionar hierarquias de dados à Tabela Dinâmica

As hierarquias de dados preenchem a Tabela Dinâmica com informações a combinar com base nas linhas e colunas. A adição das hierarquias de dados de Engradados **vendidos** no farm e engradados vendidos **no** farm fornece somas dessas figuras para cada linha e coluna.

No exemplo, **Farm** e **Type** são linhas, com as vendas de engradados como os dados.

![Uma Tabela Dinâmica mostrando o total de vendas de diferentes árvores com base no farm de onde elas vieram.](../images/excel-pivots-data-hierarchy.png)

```js
Excel.run(function (context) {
    var pivotTable = context.workbook.worksheets.getActiveWorksheet().pivotTables.getItem("Farm Sales");

    // "Farm" and "Type" are the hierarchies on which the aggregation is based.
    pivotTable.rowHierarchies.add(pivotTable.hierarchies.getItem("Farm"));
    pivotTable.rowHierarchies.add(pivotTable.hierarchies.getItem("Type"));

    // "Crates Sold at Farm" and "Crates Sold Wholesale" are the hierarchies
    // that will have their data aggregated (summed in this case).
    pivotTable.dataHierarchies.add(pivotTable.hierarchies.getItem("Crates Sold at Farm"));
    pivotTable.dataHierarchies.add(pivotTable.hierarchies.getItem("Crates Sold Wholesale"));

    return context.sync();
});
```

## <a name="pivottable-layouts-and-getting-pivoted-data"></a>Layouts de tabela dinâmica e obter dados de tabela dinâmica

Um [PivotLayout](/javascript/api/excel/excel.pivotlayout) define o posicionamento das hierarquias e seus dados. Você acessa o layout para determinar os intervalos onde os dados são armazenados.

O diagrama a seguir mostra quais chamadas de função de layout correspondem a quais intervalos da tabela dinâmica.

![Um diagrama mostrando quais seções de uma tabela dinâmica são retornadas pelas funções de intervalo de obter do layout.](../images/excel-pivots-layout-breakdown.png)

### <a name="get-data-from-the-pivottable"></a>Obter dados da tabela dinâmica

O layout define como a tabela dinâmica é exibida na planilha. Isso significa que `PivotLayout` o objeto controla os intervalos usados para elementos de tabela dinâmica. Use os intervalos fornecidos pelo layout para obter dados coletados e agregados pela Tabela Dinâmica. Em particular, use `PivotLayout.getDataBodyRange` para acessar o que a Tabela Dinâmica produz.

O código a seguir demonstra como obter a última linha dos dados da tabela dinâmica passando pelo layout (o **total geral** da soma dos engradados **vendidos** no farm e a soma das colunas de animais **vendidos no** exemplo anterior). Esses valores são somados para um total final, que é exibido na célula **E30** (fora da tabela dinâmica).

```js
Excel.run(function (context) {
    var pivotTable = context.workbook.worksheets.getActiveWorksheet().pivotTables.getItem("Farm Sales");

    // Get the totals for each data hierarchy from the layout.
    var range = pivotTable.layout.getDataBodyRange();
    var grandTotalRange = range.getLastRow();
    grandTotalRange.load("address");
    return context.sync().then(function () {
        // Sum the totals from the PivotTable data hierarchies and place them in a new range, outside of the PivotTable.
        var masterTotalRange = context.workbook.worksheets.getActiveWorksheet().getRange("E30");
        masterTotalRange.formulas = [["=SUM(" + grandTotalRange.address + ")"]];
    });
});
```

### <a name="layout-types"></a>Tipos de layout

As Tabelas Dinâmicas têm três estilos de layout: Compacto, Contorno e Tabular. Vimos o estilo compacto nos exemplos anteriores.

Os exemplos a seguir usam os estilos de contorno e tabular, respectivamente. O exemplo de código mostra como fazer o ciclo entre os diferentes layouts.

#### <a name="outline-layout"></a>Layout da estrutura de contorno

![Uma Tabela Dinâmica usando o layout da estrutura de estrutura de estrutura.](../images/excel-pivots-outline-layout.png)

#### <a name="tabular-layout"></a>Layout tabular

![Uma tabela dinâmica usando o layout tabular.](../images/excel-pivots-tabular-layout.png)

## <a name="delete-a-pivottable"></a>Excluir uma Tabela Dinâmica

As Tabelas Dinâmicas são excluídas usando seu nome.

```js
Excel.run(function (context) {
    context.workbook.worksheets.getItem("Pivot").pivotTables.getItem("Farm Sales").delete();
    return context.sync();
});
```

## <a name="filter-a-pivottable"></a>Filtrar uma Tabela Dinâmica

O método principal para filtrar dados de tabela dinâmica é com PivotFilters. As slicers oferecem um método de filtragem alternativo e menos flexível. 

[Os PivotFilters](/javascript/api/excel/excel.pivotfilters) filtram dados com [](#hierarchies) base nas quatro categorias de hierarquia de uma Tabela Dinâmica (filtros, colunas, linhas e valores). Há quatro tipos de PivotFilters, permitindo filtragem baseada em data do calendário, análise de cadeia de caracteres, comparação de números e filtragem com base em uma entrada personalizada. 

[As slicers](/javascript/api/excel/excel.slicer) podem ser aplicadas a tabelas dinâmicas e tabelas regulares do Excel. Quando aplicadas a uma tabela dinâmica, as slicers funcionam como um [PivotManualFilter](#pivotmanualfilter) e permitem a filtragem com base em uma entrada personalizada. Ao contrário de PivotFilters, as slicers têm um [componente de interface do usuário do Excel.](https://support.office.com/article/Use-slicers-to-filter-data-249f966b-a9d5-4b0f-b31a-12651785d29d) Com a `Slicer` classe, você cria esse componente da interface do usuário, gerencia a filtragem e controla sua aparência visual. 

### <a name="filter-with-pivotfilters"></a>Filtrar com PivotFilters

[Os PivotFilters](/javascript/api/excel/excel.pivotfilters) permitem filtrar dados [](#hierarchies) de Tabela Dinâmica com base nas quatro categorias de hierarquia (filtros, colunas, linhas e valores). No modelo de objeto de tabela dinâmica, `PivotFilters` são aplicados a [um PivotField](/javascript/api/excel/excel.pivotfield)e cada um pode `PivotField` ter um ou mais `PivotFilters` atribuídos. Para aplicar PivotFilters a um PivotField, o [PivotHierarchy](/javascript/api/excel/excel.pivothierarchy) correspondente do campo deve ser atribuído a uma categoria de hierarquia. 

#### <a name="types-of-pivotfilters"></a>Tipos de PivotFilters

| Tipo de filtro | Finalidade do filtro | Referência da API JavaScript do Excel |
|:--- |:--- |:--- |
| DateFilter | Filtragem baseada em data do calendário. | [PivotDateFilter](/javascript/api/excel/excel.pivotdatefilter) |
| LabelFilter | Filtragem de comparação de texto. | [PivotLabelFilter](/javascript/api/excel/excel.pivotlabelfilter) |
| ManualFilter | Filtragem de entrada personalizada. | [PivotManualFilter](/javascript/api/excel/excel.pivotmanualfilter) |
| ValueFilter | Filtragem de comparação de números. | [PivotValueFilter](/javascript/api/excel/excel.pivotvaluefilter) |

#### <a name="create-a-pivotfilter"></a>Criar um PivotFilter

Para filtrar dados de tabela dinâmica com um `Pivot*Filter` (como um `PivotDateFilter` ), aplique o filtro a um [PivotField](/javascript/api/excel/excel.pivotfield). Os quatro exemplos de código a seguir mostram como usar cada um dos quatro tipos de PivotFilters. 

##### <a name="pivotdatefilter"></a>PivotDateFilter

O primeiro exemplo de código aplica  um [PivotDateFilter](/javascript/api/excel/excel.pivotdatefilter) ao PivotField data atualizada, ocultando todos os dados anteriores a **2020-08-01**. 

> [!IMPORTANT] 
> Um `Pivot*Filter` não pode ser aplicado a um PivotField, a menos que PivotHierarchy desse campo seja atribuído a uma categoria de hierarquia. No exemplo de código a seguir, o valor deve ser adicionado à categoria da Tabela Dinâmica antes que possa `dateHierarchy` `rowHierarchies` ser usado para filtragem.

```js
Excel.run(function (context) {
    // Get the PivotTable and the date hierarchy.
    var pivotTable = context.workbook.worksheets.getActiveWorksheet().pivotTables.getItem("Farm Sales");
    var dateHierarchy = pivotTable.rowHierarchies.getItemOrNullObject("Date Updated");
    
    return context.sync().then(function () {
        // PivotFilters can only be applied to PivotHierarchies that are being used for pivoting.
        // If it's not already there, add "Date Updated" to the hierarchies.
        if (dateHierarchy.isNullObject) {
          dateHierarchy = pivotTable.rowHierarchies.add(pivotTable.hierarchies.getItem("Date Updated"));
        }

        // Apply a date filter to filter out anything logged before August.
        var filterField = dateHierarchy.fields.getItem("Date Updated");
        var dateFilter = {
          condition: Excel.DateFilterCondition.afterOrEqualTo,
          comparator: {
            date: "2020-08-01",
            specificity: Excel.FilterDatetimeSpecificity.month
          }
        };
        filterField.applyFilter({ dateFilter: dateFilter });
        
        return context.sync();
    });
});
```

> [!NOTE]
> Os três trechos de código a seguir exibem apenas trechos de filtro específicos, em vez de chamadas `Excel.run` completas.

##### <a name="pivotlabelfilter"></a>PivotLabelFilter

O segundo trecho de código demonstra como aplicar um [PivotLabelFilter](/javascript/api/excel/excel.pivotlabelfilter) a **Type** PivotField, usando a propriedade para excluir rótulos que começam com a `LabelFilterCondition.beginsWith` letra **L**. 

```js
    // Get the "Type" field.
    var filterField = pivotTable.hierarchies.getItem("Type").fields.getItem("Type");

    // Filter out any types that start with "L" ("Lemons" and "Limes" in this case).
    var filter: Excel.PivotLabelFilter = {
      condition: Excel.LabelFilterCondition.beginsWith,
      substring: "L",
      exclusive: true
    };

    // Apply the label filter to the field.
    filterField.applyFilter({ labelFilter: filter });
```

##### <a name="pivotmanualfilter"></a>PivotManualFilter

O terceiro trecho de código aplica um filtro manual com  [PivotManualFilter](/javascript/api/excel/excel.pivotmanualfilter) ao campo Classificação, filtrando dados que não incluem a classificação **orgânica**. 

```js
    // Apply a manual filter to include only a specific PivotItem (the string "Organic").
    var filterField = classHierarchy.fields.getItem("Classification");
    var manualFilter = { selectedItems: ["Organic"] };
    filterField.applyFilter({ manualFilter: manualFilter });
```

##### <a name="pivotvaluefilter"></a>PivotValueFilter

Para comparar números, use um filtro de valor com [PivotValueFilter](/javascript/api/excel/excel.pivotvaluefilter), conforme mostrado no trecho de código final. A comparação dos dados no Farm PivotField com os dados no PivotField de Campos Dinâmicas de Engradados vendidos, incluindo apenas farms cuja soma de engradados vendidos excede o valor `PivotValueFilter` **500**.   

```js
    // Get the "Farm" field.
    var filterField = pivotTable.hierarchies.getItem("Farm").fields.getItem("Farm");
    
    // Filter to only include rows with more than 500 wholesale crates sold.
    var filter: Excel.PivotValueFilter = {
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
Excel.run(function (context) {
    // Get the PivotTable.
    var pivotTable = context.workbook.worksheets.getActiveWorksheet().pivotTables.getItem("Farm Sales");
    pivotTable.hierarchies.load("name");
    
    return context.sync().then(function () {
        // Clear the filters on each PivotField.
        pivotTable.hierarchies.items.forEach(function (hierarchy) {
          hierarchy.fields.getItem(hierarchy.name).clearAllFilters();
        });
        return context.sync();
    });
});
```

### <a name="filter-with-slicers"></a>Filtrar com slicers

[As slicers](/javascript/api/excel/excel.slicer) permitem que os dados sejam filtrados de uma tabela ou tabela dinâmica do Excel. Uma slicer usa valores de uma coluna especificada ou PivotField para filtrar as linhas correspondentes. Esses valores são armazenados [como objetos SlicerItem](/javascript/api/excel/excel.sliceritem) no `Slicer` objeto . Seu complemento pode ajustar esses filtros, assim como os usuários[(por meio da interface do usuário do Excel).](https://support.office.com/article/Use-slicers-to-filter-data-249f966b-a9d5-4b0f-b31a-12651785d29d) A slicer fica na parte superior da planilha na camada de desenho, conforme mostrado na captura de tela a seguir.

![Dados de filtragem de uma slicer em uma tabela dinâmica.](../images/excel-slicer.png)

> [!NOTE]
> As técnicas descritas nesta seção concentram-se em como usar as slicers conectadas a Tabelas Dinâmicas. As mesmas técnicas também se aplicam ao uso de slicers conectados a tabelas.

#### <a name="create-a-slicer"></a>Criar uma slicer

Você pode criar uma slicer em uma pasta de trabalho ou planilha usando o `Workbook.slicers.add` método ou `Worksheet.slicers.add` método. Isso adiciona uma slicer à [SlicerCollection](/javascript/api/excel/excel.slicercollection) do objeto `Workbook` `Worksheet` ou especificado. O `SlicerCollection.add` método tem três parâmetros:

- `slicerSource`: a fonte de dados na qual a nova slicer se baseia. Pode ser uma cadeia de caracteres , ou `PivotTable` `Table` representando o nome ou a ID de um `PivotTable` ou `Table` .
- `sourceField`: O campo na fonte de dados pelo qual filtrar. Pode ser uma `PivotField` cadeia de `TableColumn` caracteres , ou representando o nome ou a ID de um `PivotField` ou `TableColumn` .
- `slicerDestination`: a planilha onde a nova slicer será criada. Pode ser um `Worksheet` objeto ou o nome ou A ID de um `Worksheet` . Esse parâmetro é desnecessário quando é `SlicerCollection` acessado por meio de `Worksheet.slicers` . Nesse caso, a planilha da coleção é usada como destino.

O exemplo de código a seguir adiciona uma nova slicer à **planilha** Dinâmica. A origem da slicer é a Tabela Dinâmica **de Vendas** do Farm e filtra usando os **dados Type.** A slicer também é chamada de **Slicer de** Fatia para referência futura.

```js
Excel.run(function (context) {
    var sheet = context.workbook.worksheets.getItem("Pivot");
    var slicer = sheet.slicers.add(
        "Farm Sales" /* The slicer data source. For PivotTables, this can be the PivotTable object reference or name. */,
        "Type" /* The field in the data to filter by. For PivotTables, this can be a PivotField object reference or ID. */
    );
    slicer.name = "Fruit Slicer";
    return context.sync();
});
```

#### <a name="filter-items-with-a-slicer"></a>Filtrar itens com uma slicer

A slicer filtra a Tabela Dinâmica com itens do `sourceField` . O `Slicer.selectItems` método define os itens que permanecem na slicer. Esses itens são passados para o método como um `string[]` , representando as chaves dos itens. Todas as linhas que contêm esses itens permanecem na agregação da Tabela Dinâmica. Chamadas subsequentes `selectItems` para definir a lista com as chaves especificadas nessas chamadas.

> [!NOTE]
> Se `Slicer.selectItems` for passado um item que não está na fonte de dados, será lançado `InvalidArgument` um erro. O conteúdo pode ser verificado por meio da `Slicer.slicerItems` propriedade, que é uma [SlicerItemCollection](/javascript/api/excel/excel.sliceritemcollection).

O exemplo de código a seguir mostra três itens sendo selecionados para a slicer: **Limão**, **Verde-limão** e **Laranja**.

```js
Excel.run(function (context) {
    var slicer = context.workbook.slicers.getItem("Fruit Slicer");
    // Anything other than the following three values will be filtered out of the PivotTable for display and aggregation.
    slicer.selectItems(["Lemon", "Lime", "Orange"]);
    return context.sync();
});
```

Para remover todos os filtros da slicer, use o `Slicer.clearFilters` método, conforme mostrado no exemplo a seguir.

```js
Excel.run(function (context) {
    var slicer = context.workbook.slicers.getItem("Fruit Slicer");
    slicer.clearFilters();
    return context.sync();
});
```

#### <a name="style-and-format-a-slicer"></a>Estilo e formatar uma slicer

Você pode ajustar as configurações de exibição de uma slicer por meio de `Slicer` propriedades. O exemplo de código a seguir define o estilo como **SlicerStyleLight6**, define o texto na parte superior da slicer como **Tipos** de Árvores , coloca a slicer na posição **(395, 15)** na camada de desenho e define o tamanho da slicer como **135x150** pixels.

```js
Excel.run(function (context) {
    var slicer = context.workbook.slicers.getItem("Fruit Slicer");
    slicer.caption = "Fruit Types";
    slicer.left = 395;
    slicer.top = 15;
    slicer.height = 135;
    slicer.width = 150;
    slicer.style = "SlicerStyleLight6";
    return context.sync();
});
```

#### <a name="delete-a-slicer"></a>Excluir uma slicer

Para excluir uma slicer, chame o `Slicer.delete` método. O exemplo de código a seguir exclui a primeira fatia da planilha atual.

```js
Excel.run(function (context) {
    var sheet = context.workbook.worksheets.getActiveWorksheet();
    sheet.slicers.getItemAt(0).delete();
    return context.sync();
});
```

## <a name="change-aggregation-function"></a>Alterar a função de agregação

Hierarquias de dados têm seus valores agregados. Para conjuntos de dados de números, essa é uma soma por padrão. A `summarizeBy` propriedade define esse comportamento com base em um tipo [AggregationFunction.](/javascript/api/excel/excel.aggregationfunction)

Os tipos de função de agregação atualmente suportados são `Sum` `Count` , `Average` e `Max` `Min` `Product` `CountNumbers` `StandardDeviation` `StandardDeviationP` `Variance` `VarianceP` `Automatic` (o padrão).

Os exemplos de código a seguir altera a agregação para as médias dos dados.

```js
Excel.run(function (context) {
    var pivotTable = context.workbook.worksheets.getActiveWorksheet().pivotTables.getItem("Farm Sales");
    pivotTable.dataHierarchies.load("no-properties-needed");
    return context.sync().then(function() {

        // Change the aggregation from the default sum to an average of all the values in the hierarchy.
        pivotTable.dataHierarchies.items[0].summarizeBy = Excel.AggregationFunction.average;
        pivotTable.dataHierarchies.items[1].summarizeBy = Excel.AggregationFunction.average;
        return context.sync();
    });
});
```

## <a name="change-calculations-with-a-showasrule"></a>Alterar cálculos com um ShowAsRule

As Tabelas Dinâmicas, por padrão, agregam os dados de suas hierarquias de linha e coluna de forma independente. Um [ShowAsRule altera](/javascript/api/excel/excel.showasrule) a hierarquia de dados para valores de saída com base em outros itens na tabela dinâmica.

O `ShowAsRule` objeto tem três propriedades:

- `calculation`: o tipo de cálculo relativo a ser aplicado à hierarquia de dados (o padrão é `none` ).
- `baseField`: [PivotField](/javascript/api/excel/excel.pivotfield) na hierarquia que contém os dados base antes da aplicação do cálculo. Como as Tabelas Dinâmicas do Excel têm um mapeamento de hierarquia para um de campo, você usará o mesmo nome para acessar a hierarquia e o campo.
- `baseItem`: o [PivotItem](/javascript/api/excel/excel.pivotitem) individual comparado com os valores dos campos base com base no tipo de cálculo. Nem todos os cálculos exigem esse campo.

O exemplo a seguir define o cálculo na Soma de Caixas **vendidas** na hierarquia de dados do farm como uma porcentagem do total da coluna.
Ainda queremos que a granularidade se estenda até o nível do tipo de resultado, portanto, vamos usar a hierarquia de linhas **Type** e seu campo subjacente.
O exemplo também tem **Farm** como a primeira hierarquia de linhas, portanto, as entradas totais do farm exibem a porcentagem pela qual cada farm também é responsável pela produção.

![Uma Tabela Dinâmica mostrando as porcentagens de vendas de árvores relativas ao total geral para farms individuais e tipos de árvores individuais em cada farm.](../images/excel-pivots-showas-percentage.png)

```js
Excel.run(function (context) {
    var pivotTable = context.workbook.worksheets.getActiveWorksheet().pivotTables.getItem("Farm Sales");
    var farmDataHierarchy = pivotTable.dataHierarchies.getItem("Sum of Crates Sold at Farm");

    farmDataHierarchy.load("showAs");
    return context.sync().then(function () {

        // Show the crates of each fruit type sold at the farm as a percentage of the column's total.
        var farmShowAs = farmDataHierarchy.showAs;
        farmShowAs.calculation = Excel.ShowAsCalculation.percentOfColumnTotal;
        farmShowAs.baseField = pivotTable.rowHierarchies.getItem("Type").fields.getItem("Type");
        farmDataHierarchy.showAs = farmShowAs;
        farmDataHierarchy.name = "Percentage of Total Farm Sales";
    });
});
```

O exemplo anterior definiu o cálculo para a coluna, em relação ao campo de uma hierarquia de linhas individual. Quando o cálculo estiver relacionado a um item individual, use a `baseItem` propriedade.

O exemplo a seguir mostra o `differenceFrom` cálculo. Ele exibe a diferença das entradas da hierarquia de dados de vendas do farm em relação às de **farms A.**
O farm é, portanto, vemos as diferenças entre os outros farms, bem como os detalhamentos de cada tipo de like fruit ( Type também é uma hierarquia de linhas `baseField` neste exemplo). 

![Uma Tabela Dinâmica mostrando as diferenças de vendas de árvores entre "Farms A" e os outros. Isso mostra a diferença no total de vendas de árvores dos farms e nas vendas de tipos de árvores. Se "A Farms" não tiver vendido um tipo específico de resultado, "#N/A" será exibido.](../images/excel-pivots-showas-differencefrom.png)

```js
Excel.run(function (context) {
    var pivotTable = context.workbook.worksheets.getActiveWorksheet().pivotTables.getItem("Farm Sales");
    var farmDataHierarchy = pivotTable.dataHierarchies.getItem("Sum of Crates Sold at Farm");

    farmDataHierarchy.load("showAs");
    return context.sync().then(function () {
        // Show the difference between crate sales of the "A Farms" and the other farms.
        // This difference is both aggregated and shown for individual fruit types (where applicable).
        var farmShowAs = farmDataHierarchy.showAs;
        farmShowAs.calculation = Excel.ShowAsCalculation.differenceFrom;
        farmShowAs.baseField = pivotTable.rowHierarchies.getItem("Farm").fields.getItem("Farm");
        farmShowAs.baseItem = pivotTable.rowHierarchies.getItem("Farm").fields.getItem("Farm").items.getItem("A Farms");
        farmDataHierarchy.showAs = farmShowAs;
        farmDataHierarchy.name = "Difference from A Farms";
    });
});
```

## <a name="change-hierarchy-names"></a>Alterar nomes de hierarquia

Os campos de hierarquia são editáveis. O código a seguir demonstra como alterar os nomes exibidos de duas hierarquias de dados.

```js
Excel.run(function (context) {
    var dataHierarchies = context.workbook.worksheets.getActiveWorksheet()
        .pivotTables.getItem("Farm Sales").dataHierarchies;
    dataHierarchies.load("no-properties-needed");
    return context.sync().then(function () {
        // changing the displayed names of these entries
        dataHierarchies.items[0].name = "Farm Sales";
        dataHierarchies.items[1].name = "Wholesale";
    });
});
```

## <a name="see-also"></a>Confira também

- [Modelo de objeto JavaScript do Excel em Suplementos do Office](excel-add-ins-core-concepts.md)
- [Referência da API JavaScript do Excel](/javascript/api/excel)
