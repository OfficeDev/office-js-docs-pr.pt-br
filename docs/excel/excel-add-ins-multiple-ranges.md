---
title: Trabalhar simultaneamente com vários intervalos em suplementos do Excel
description: Saiba como a biblioteca JavaScript do Excel permite que o suplemento realize operações e defina propriedades em vários intervalos simultaneamente.
ms.date: 04/30/2019
localization_priority: Normal
ms.openlocfilehash: 97481b4b8ab76f7bbc5bd10378d4cc6512bc7b6a
ms.sourcegitcommit: fa4e81fcf41b1c39d5516edf078f3ffdbd4a3997
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 03/17/2020
ms.locfileid: "42717065"
---
# <a name="work-with-multiple-ranges-simultaneously-in-excel-add-ins"></a>Trabalhar simultaneamente com vários intervalos em suplementos do Excel

A biblioteca de JavaScript do Excel permite que o suplemento realize operações e defina propriedades, em vários intervalos simultaneamente. Os intervalos não precisam ser contíguos. Além de tornar seu código mais simples, essa maneira de definir uma propriedade é executada muito mais rapidamente do que definir a mesma propriedade individualmente para cada um dos intervalos.

## <a name="rangeareas"></a>RangeAreas

Um conjunto de intervalos (possivelmente não contíguos) é representado por um objeto [RangeAreas](/javascript/api/excel/excel.rangeareas) . Possui propriedades e métodos semelhantes ao tipo `Range` (muitos com os mesmos nomes ou semelhantes), mas foram feitos ajustes para:

- Os tipos de dados para propriedades e o comportamento dos setters e getters.
- Os tipos de dados dos parâmetros do método e os comportamentos do método.
- Os tipos de dados de forma retornam valores.

Alguns exemplos:

- `RangeAreas` tem uma propriedade `address` que retorna uma cadeia de caracteres delimitada por vírgula de intervalo de endereços, em vez de apenas um endereço como na propriedade`Range.address`.
- `RangeAreas` tem uma propriedade `dataValidation` que retorna um objeto `DataValidation` que representa a validação de dados de todos os intervalos em`RangeAreas`, se for consistente. A propriedade é `null` se objetos idênticos `DataValidation` não forem aplicados a todos os intervalos em `RangeAreas`. Esse é um princípio geral, mas não universal com o objeto `RangeAreas`: *se uma propriedade não têm valores consistentes em todos os todos os intervalos em `RangeAreas`, então será `null`.* Ver [ler as propriedades de RangeAreas](#read-properties-of-rangeareas) para mais informações e algumas exceções.
- `RangeAreas.cellCount` é o número total de células em todos os intervalos no `RangeAreas`.
- `RangeAreas.calculate` recalcula as células de todos os intervalos no `RangeAreas`.
- `RangeAreas.getEntireColumn` e `RangeAreas.getEntireRow` retornar outra `RangeAreas` objeto que representa todas as colunas (ou linhas) em todos os intervalos no `RangeAreas`. Por exemplo, se `RangeAreas` representa "A1: C4" e "F14:L15" em seguida, `RangeAreas.getEntireColumn` retorna um objeto `RangeAreas` que representa "A:C" e "F:L".
- `RangeAreas.copyFrom` pode ter o parâmetro `Range` ou `RangeAreas` que representam os intervalos de origem da operação de cópia.

#### <a name="complete-list-of-range-members-that-are-also-available-on-rangeareas"></a>Lista completa de membros do intervalo que também estão disponíveis em RangeAreas

##### <a name="properties"></a>Propriedades

Familiarize-se com as [Propriedades de leitura do RangeAreas](#read-properties-of-rangeareas) antes de escrever o código que lê as propriedades listadas. Existem sutilezas para o que é retornado.

- `address`
- `addressLocal`
- `cellCount`
- `conditionalFormats`
- `context`
- `dataValidation`
- `format`
- `isEntireColumn`
- `isEntireRow`
- `style`
- `worksheet`

##### <a name="methods"></a>Métodos

- `calculate()`
- `clear()`
- `convertDataTypeToText()`
- `convertToLinkedDataType()`
- `copyFrom()`
- `getEntireColumn()`
- `getEntireRow()`
- `getIntersection()`
- `getIntersectionOrNullObject()`
- `getOffsetRange()`(nomeado `getOffsetRangeAreas` no `RangeAreas` objeto)
- `getSpecialCells()`
- `getSpecialCellsOrNullObject()`
- `getTables()`
- `getUsedRange()`(nomeado `getUsedRangeAreas` no `RangeAreas` objeto)
- `getUsedRangeOrNullObject()`(nomeado `getUsedRangeAreasOrNullObject` no `RangeAreas` objeto)
- `load()`
- `set()`
- `setDirty()`
- `toJSON()`
- `track()`
- `untrack()`

### <a name="rangearea-specific-properties-and-methods"></a>Métodos e propriedades específicos do RangeArea

O tipo `RangeAreas` tem alguns métodos e propriedades que não estão no objeto `Range`. Esta é a seleção deles:

- `areas`: O objeto `RangeCollection` que contém todos os intervalos representados pelo objeto `RangeAreas`. O objeto `RangeCollection` também é novidade e é semelhante a outros objetos do conjunto do Excel. É uma propriedade `items` que é uma matriz de objetos `Range` que representam os intervalos.
- `areaCount`: O número total de intervalos em `RangeAreas`.
- `getOffsetRangeAreas`: Funciona como [Range.getOffsetRange](/javascript/api/excel/excel.range#getoffsetrange-rowoffset--columnoffset-), exceto pelo fato de que o `RangeAreas` é retornado e contém os intervalos que são todos os deslocamentos de um dos intervalos do `RangeAreas` original.

## <a name="create-rangeareas"></a>Criar RangeAreas

Você pode criar o objeto`RangeAreas` de duas maneiras básicas:

- Ligue `Worksheet.getRanges()` e encaminhe-o em uma cadeia de caracteres com endereços de intervalo separado por vírgula. Se algum intervalo que você deseja incluir tiver sido feito em um [NamedItem](/javascript/api/excel/excel.nameditem), você poderá incluir o nome, em vez do endereço, cadeia de caracteres.
- Chamar `Workbook.getSelectedRanges()`. Esse método retornará um `RangeAreas` representando todos os intervalos selecionados na planilha ativa no momento.

Quando você tiver um objeto `RangeAreas`, você pode criar outros usando os métodos de objeto que retornam `RangeAreas` como `getOffsetRangeAreas` e `getIntersection`.

> [!NOTE]
> É possível adicionar diretamente intervalos adicionais para um objeto `RangeAreas`. Por exemplo, o conjunto `RangeAreas.areas` não tem um método`add`.

> [!WARNING]
> Tente adicionar ou excluir membros diretamente à matriz`RangeAreas.areas.items`. Isso levará a um comportamento indesejável no seu código. Por exemplo, é possível enviar um objeto adicional `Range` para a matriz, mas isso causará erros porque as propriedades e métodos `RangeAreas` se comportam como se o novo item não estivesse ali. Por exemplo, a propriedade `areaCount` não inclui intervalos transferidos dessa maneira e o `RangeAreas.getItemAt(index)` gera um erro se `index` for maior que `areasCount-1`. Da mesma forma, excluir um objeto `Range` na matriz `RangeAreas.areas.items` obtendo uma referência a ele e chamando seu método `Range.delete` causa bugs: embora o `Range`objeto* seja *excluído, as propriedades e métodos do objeto pai `RangeAreas` se comportam ou tentam se comportar, como se ele ainda existisse. Por exemplo, se o seu código chamar `RangeAreas.calculate`, o Office tentará calcular o intervalo, mas haverá erro porque o objeto de intervalo desapareceu.

## <a name="set-properties-on-multiple-ranges"></a>Definir as propriedades em vários intervalos

Definir uma propriedade em um `RangeAreas` objeto define a propriedade correspondente em todos os intervalos no conjunto `RangeAreas.areas`.

A seguir, um exemplo de configuração de uma propriedade em vários intervalos. A função realça os intervalos **F3:F5** e **H3:H5**.

```js
Excel.run(function (context) {
    var sheet = context.workbook.worksheets.getActiveWorksheet();
    var rangeAreas = sheet.getRanges("F3:F5, H3:H5");
    rangeAreas.format.fill.color = "pink";

    return context.sync();
})
```

Este exemplo se aplica a cenários nos quais você pode codificar os endereços de intervalo para os quais você passa para `getRanges` ou facilmente calculá-los no tempo de execução. Alguns dos cenários em que isso pode ser verdadeiro incluem:

- O código é executado no contexto de um modelo conhecido.
- O código é executado no contexto de dados importados, em que o esquema dos dados é conhecido.

## <a name="get-special-cells-from-multiple-ranges"></a>Obter células especiais de vários intervalos

As `getSpecialCells` e `getSpecialCellsOrNullObject` métodos no `RangeAreas` objeto funciona analogamente para métodos de mesmo nome no `Range` objeto. Esses métodos retornam as células com característica especificada de todos os intervalos no `RangeAreas.areas` conjunto. Confira a seção [Localizar células especiais em um intervalo](excel-add-ins-ranges-advanced.md#find-special-cells-within-a-range) para saber mais sobre células especiais.

Ao chamar as `getSpecialCells` ou `getSpecialCellsOrNullObject` método em um `RangeAreas` objeto:

- Se você passar `Excel.SpecialCellType.sameConditionalFormat` como o primeiro parâmetro, o método retorna todas as células com a mesma formatação condicional que a célula superior esquerda do primeiro intervalo no `RangeAreas.areas` conjunto.
- Se você passar `Excel.SpecialCellType.sameDataValidation` como o primeiro parâmetro, o método retorna todas as células com a regra de validação de dados que a célula superior esquerda do primeiro intervalo no `RangeAreas.areas` conjunto.

## <a name="read-properties-of-rangeareas"></a>Ler propriedades de RangeAreas

A leitura de valores de propriedade `RangeAreas` requer cuidados, porque uma determinada propriedade pode ter valores diferentes para intervalos diferentes dentro de`RangeAreas`. A regra geral é que, se um valor consistente *puder* ser retornado, ele será retornado. Por exemplo, no código a seguir, O código RGB para pink (`#FFC0CB`) e `true` será registrado no console porque ambos os intervalos no objeto `RangeAreas` têm um preenchimento rosa e ambos são colunas inteiras.

```js
Excel.run(function (context) {
    var sheet = context.workbook.worksheets.getActiveWorksheet();

    // The ranges are the F column and the H column.
    var rangeAreas = sheet.getRanges("F:F, H:H");  
    rangeAreas.format.fill.color = "pink";

    rangeAreas.load("format/fill/color, isEntireColumn");

    return context.sync()
        .then(function () {
            console.log(rangeAreas.format.fill.color); // #FFC0CB
            console.log(rangeAreas.isEntireColumn); // true
        })
        .then(context.sync);
})
```

As coisas ficam mais complicadas quando a consistência não é possível. O comportamento das propriedades `RangeAreas` seguem estes três princípios de três:

- Uma propriedade booliana de um `RangeAreas`retorno de objeto `false`, a menos que a propriedade seja verdadeira para todos os intervalos de membro.
- Propriedades não boolianas, com exceção da propriedade `address`, retornam `null`, a menos que a propriedade correspondente em todos os intervalos de membros tenha o mesmo valor.
- A propriedade `address` retorna uma cadeia de caracteres delimitada por vírgulas dos endereços e intervalos dos membros.

Por exemplo, o código a seguir cria um `RangeAreas` no qual apenas um intervalo é uma coluna inteira e apenas um é preenchido com rosa. O console mostrará `null` para a cor de preenchimento `false` para a propriedade `isEntireRow` e "Planilha1! F3:F5, Planilha1! H:H"(supondo que o nome da planilha  seja "Planilha1") para a propriedade`address`.

```js
Excel.run(function (context) {
    var sheet = context.workbook.worksheets.getActiveWorksheet();
    var rangeAreas = sheet.getRanges("F3:F5, H:H");

    var pinkColumnRange = sheet.getRange("H:H");
    pinkColumnRange.format.fill.color = "pink";

    rangeAreas.load("format/fill/color, isEntireColumn, address");

    return context.sync()
        .then(function () {
            console.log(rangeAreas.format.fill.color); // null
            console.log(rangeAreas.isEntireColumn); // false
            console.log(rangeAreas.address); // "Sheet1!F3:F5, Sheet1!H:H"
        })
        .then(context.sync);
})
```

## <a name="see-also"></a>Confira também

- [Conceitos fundamentais de programação com a API JavaScript do Excel](../reference/overview/excel-add-ins-reference-overview.md)
- [Trabalhe com intervalos usando a API JavaScript do Excel (fundamental)](excel-add-ins-ranges.md)
- [Trabalhe com intervalos usando a API JavaScript do Excel (avançado)](excel-add-ins-ranges-advanced.md)
