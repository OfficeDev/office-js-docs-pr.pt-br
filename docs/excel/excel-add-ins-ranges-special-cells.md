---
title: Encontre células especiais em um intervalo usando a api javascript Excel javascript
description: Saiba como usar a EXCEL JavaScript para encontrar células especiais, como células com fórmulas, erros ou números.
ms.date: 02/17/2022
ms.prod: excel
ms.localizationpriority: medium
ms.openlocfilehash: 2705f71b5b253801493b07143b8320a1f69062e9
ms.sourcegitcommit: 968d637defe816449a797aefd930872229214898
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 03/23/2022
ms.locfileid: "63744793"
---
# <a name="find-special-cells-within-a-range-using-the-excel-javascript-api"></a>Encontre células especiais em um intervalo usando a api javascript Excel javascript

Este artigo fornece exemplos de código que encontram células especiais dentro de um intervalo usando Excel API JavaScript. Para ver a lista completa de propriedades e métodos compatíveis `Range` com o objeto, [consulte Excel. Classe Range](/javascript/api/excel/excel.range).

## <a name="find-ranges-with-special-cells"></a>Encontrar intervalos com células especiais

Os [métodos Range.getSpecialCells](/javascript/api/excel/excel.range#excel-excel-range-getspecialcells-member(1)) e [Range.getSpecialCellsOrNullObject](/javascript/api/excel/excel.range#excel-excel-range-getspecialcellsornullobject-member(1)) encontram intervalos com base nas características de suas células e nos tipos de valores de suas células. Os dois métodos retornam `RangeAreas` objetos. Aqui estão as assinaturas dos métodos do arquivo de tipos de dados TypeScript:

```typescript
getSpecialCells(cellType: Excel.SpecialCellType, cellValueType?: Excel.SpecialCellValueType): Excel.RangeAreas;
```

```typescript
getSpecialCellsOrNullObject(cellType: Excel.SpecialCellType, cellValueType?: Excel.SpecialCellValueType): Excel.RangeAreas;
```

O exemplo de código a seguir usa o `getSpecialCells` método para encontrar todas as células com fórmulas. Sobre este código, observe:

- Ele limita a parte da planilha que precisa ser pesquisada chamando primeiro `Worksheet.getUsedRange` e chamando `getSpecialCells` para apenas esse intervalo.
- O método `getSpecialCells` retorna um objeto `RangeAreas`, então todas as células com fórmulas serão coloridas de rosa, mesmo que não sejam todas contíguas.

```js
await Excel.run(async (context) => {
    let sheet = context.workbook.worksheets.getActiveWorksheet();
    let usedRange = sheet.getUsedRange();
    let formulaRanges = usedRange.getSpecialCells(Excel.SpecialCellType.formulas);
    formulaRanges.format.fill.color = "pink";

    await context.sync();
});
```

Se nenhuma célula com característica destino existe no intervalo, `getSpecialCells` exibe um erro **ItemNotFound**. Isso desvia o fluxo de controle para um `catch` bloco, se houver um. Se não houver um bloco `catch` , o erro interromperá o método.

Se você espera que células com característica direcionada sempre deveriam existir, provavelmente desejará o código para gerar um erro se as células não estiverem lá. Se for um cenário válido que não há uma ou mais células correspondentes, o código deve verificar se há essa possibilidade e tratar normalmente sem enviar um erro. Você pode obter esse comportamento com o `getSpecialCellsOrNullObject` método e sua propriedade retornada `isNullObject`. O exemplo de código a seguir usa esse padrão. Sobre este código, observe:

- O `getSpecialCellsOrNullObject` método sempre retorna um objeto proxy, portanto, nunca está `null` no sentido javaScript comum. Mas se nenhuma célula de correspondência for encontrada, as propriedades do objeto`isNullObject` serão definida como `true`.
- Ele chama `context.sync` *antes* de testar a propriedade`isNullObject`. Esse é um requisito com todos os métodos e propriedades `*OrNullObject`, pois sempre terá que carregar e sincronizar as propriedades na ordem para lê-la. No entanto, não é necessário carregar *explicitamente* a `isNullObject` propriedade. Ele é carregado automaticamente pelo `context.sync` mesmo se `load` não for chamado no objeto. Para obter mais informações, consulte [\*Métodos e propriedades OrNullObject](../develop/application-specific-api-model.md#ornullobject-methods-and-properties).
- Você pode testar esse código selecionando primeiro um intervalo sem células de fórmula e executando-o. Selecione um intervalo que tem pelo menos uma célula com uma fórmula e execute novamente.

```js
await Excel.run(async (context) => {
    let range = context.workbook.getSelectedRange();
    let formulaRanges = range.getSpecialCellsOrNullObject(Excel.SpecialCellType.formulas);
    await context.sync();
        
    if (formulaRanges.isNullObject) {
        console.log("No cells have formulas");
    }
    else {
        formulaRanges.format.fill.color = "pink";
    }
    
    await context.sync();
});
```

Para simplificar, todos os outros exemplos de código neste artigo usam o `getSpecialCells` método em vez de  `getSpecialCellsOrNullObject`.

## <a name="narrow-the-target-cells-with-cell-value-types"></a>Restrinja as células de destino com tipos de valor de célula

As `Range.getSpecialCells()` e `Range.getSpecialCellsOrNullObject()` métodos aceitam um segundo parâmetro opcional usado para restringir ainda mais as células de destino. Este segundo parâmetro é uma `Excel.SpecialCellValueType` você usar para especificar que você quer apenas células que contêm determinados tipos de valores.

> [!NOTE]
> O `Excel.SpecialCellValueType` parâmetro só pode ser usado se a `Excel.SpecialCellType` está `Excel.SpecialCellType.formulas` ou `Excel.SpecialCellType.constants`.

### <a name="test-for-a-single-cell-value-type"></a>Teste para um tipo de valor da célula única

O `Excel.SpecialCellValueType` enumeração com esses quatro tipos básicos (além dos outros valores combinados descritos nesta seção posterior):

- `Excel.SpecialCellValueType.errors`
- `Excel.SpecialCellValueType.logical` (ou seja, booliano)
- `Excel.SpecialCellValueType.numbers`
- `Excel.SpecialCellValueType.text`

O exemplo de código a seguir localiza células especiais que são constantes numéricas e colore essas células rosa. Sobre este código, observe:

- Ele só realça células que têm um valor de número literal. Ele não realça células que têm uma fórmula (mesmo que o resultado seja um número) ou um booleano, texto ou células de estado de erro.
- Para testar o código, certifique-se de que a planilha tenha algumas células com valores numéricos literais, algumas com outros tipos de valores literais e algumas com fórmulas.

```js
await Excel.run(async (context) => {
    let sheet = context.workbook.worksheets.getActiveWorksheet();
    let usedRange = sheet.getUsedRange();
    let constantNumberRanges = usedRange.getSpecialCells(
        Excel.SpecialCellType.constants,
        Excel.SpecialCellValueType.numbers);
    constantNumberRanges.format.fill.color = "pink";

    await context.sync();
});
```

### <a name="test-for-multiple-cell-value-types"></a>Teste para vários tipos de valores de célula

Às vezes, você precisa operar com mais de um tipo de valor de célula, como todas as células com valor de texto e com valor booliano ("Lógico"). (`Excel.SpecialCellValueType.logical`). O `Excel.SpecialCellValueType` enumeração tem valores com tipos combinado. Por exemplo, `Excel.SpecialCellValueType.logicalText` segmentará todas as células boolianas e todos os valores de texto. `Excel.SpecialCellValueType.all` é o valor padrão, que não limita os tipos de valor da célula retornados. O exemplo de código a seguir colore todas as células com fórmulas que produzem número ou valor booleano.

```js
await Excel.run(async (context) => {
    let sheet = context.workbook.worksheets.getActiveWorksheet();
    let usedRange = sheet.getUsedRange();
    let formulaLogicalNumberRanges = usedRange.getSpecialCells(
        Excel.SpecialCellType.formulas,
        Excel.SpecialCellValueType.logicalNumbers);
    formulaLogicalNumberRanges.format.fill.color = "pink";

    await context.sync();
});
```

## <a name="see-also"></a>Confira também

- [Modelo de objeto JavaScript do Excel em Suplementos do Office](excel-add-ins-core-concepts.md)
- [Trabalhar com células usando a EXCEL JavaScript](excel-add-ins-cells.md)
- [Encontre uma cadeia de caracteres usando Excel API JavaScript](excel-add-ins-ranges-string-match.md)
- [Trabalhar simultaneamente com vários intervalos em suplementos do Excel](excel-add-ins-multiple-ranges.md)
