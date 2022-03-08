---
title: Trabalhar com precedentes de fórmula e dependentes usando a API JavaScript Excel javascript
description: Saiba como usar a API JavaScript Excel para recuperar precedentes e dependentes da fórmula.
ms.date: 02/17/2022
ms.prod: excel
ms.localizationpriority: medium
ms.openlocfilehash: 8e401ea6dfe285a56fe0da3d250222a6e016b24c
ms.sourcegitcommit: 7b6ee73fa70b8e0ff45c68675dd26dd7a7b8c3e9
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 03/08/2022
ms.locfileid: "63340698"
---
# <a name="get-formula-precedents-and-dependents-using-the-excel-javascript-api"></a>Obter precedentes de fórmula e dependentes usando a API JavaScript Excel javascript

Excel fórmulas geralmente se referem a outras células. Essas referências entre células são conhecidas como "precedentes" e "dependentes". Um precedente é uma célula que fornece dados a uma fórmula. Um dependente é uma célula que contém uma fórmula que se refere a outras células. Para saber mais sobre os Excel relacionados às relações entre células, consulte Exibir as relações entre [fórmulas e células](https://support.microsoft.com/office/a59bef2b-3701-46bf-8ff1-d3518771d507).

Uma célula precedente pode ter suas próprias células precedentes. Cada célula precedente nessa cadeia de precedentes ainda é um precedente da célula original. A mesma relação existe para dependentes. Qualquer célula afetada por outra célula é dependente dessa célula. Um "precedente direto" é o primeiro grupo de células anterior nesta sequência, semelhante ao conceito de pais em uma relação pai-filho. Um "dependente direto" é o primeiro grupo dependente de células em uma sequência, semelhante a filhos em uma relação pai-filho.

Este artigo fornece exemplos de código que recuperam precedentes e dependentes de fórmulas usando Excel API JavaScript. Para ver a lista completa de propriedades e `Range` métodos que o objeto oferece suporte, consulte [Range Object (API JavaScript para Excel)](/javascript/api/excel/excel.range).

## <a name="get-the-precedents-of-a-formula"></a>Obter os precedentes de uma fórmula

Localize as células precedentes de uma fórmula [com Range.getPrecedents](/javascript/api/excel/excel.range#excel-excel-range-getprecedents-member(1)). `Range.getPrecedents` retorna um `WorkbookRangeAreas` objeto. Este objeto contém os endereços de todos os precedentes na workbook. Ele tem um objeto separado `RangeAreas` para cada planilha que contém pelo menos um precedente de fórmula. Para saber mais sobre o objeto`RangeAreas`, consulte [Work with multiple ranges simultaneously in Excel add-ins](excel-add-ins-multiple-ranges.md).

Para localizar apenas as células precedentes diretas de uma fórmula, use [Range.getDirectPrecedents](/javascript/api/excel/excel.range#excel-excel-range-getdirectprecedents-member(1)). `Range.getDirectPrecedents` funciona como `Range.getPrecedents` e retorna um `WorkbookRangeAreas` objeto que contém os endereços de precedentes diretos.

A captura de tela a seguir mostra o resultado da seleção do botão **Rastrear Precedentes** na interface Excel interface do usuário. Este botão desenha uma seta de células precedentes para a célula selecionada. A célula selecionada, **E3**, contém a fórmula "=C3 * D3", portanto **, C3** e **D3** são células precedentes. Ao contrário do Excel da interface do usuário, `getPrecedents` os `getDirectPrecedents` métodos e não desenham setas.

![Células precedentes de rastreamento de seta Excel interface do usuário.](../images/excel-ranges-trace-precedents.png)

> [!IMPORTANT]
> Os `getPrecedents` métodos e `getDirectPrecedents` não recuperam células precedentes entre as guias de trabalho.

O exemplo de código a seguir mostra como trabalhar com os `Range.getPrecedents` métodos e `Range.getDirectPrecedents` . O exemplo obtém os precedentes do intervalo ativo e altera a cor de plano de fundo dessas células precedentes. A cor do plano de fundo das células precedentes diretas é definida como amarelo e a cor de plano de fundo das outras células precedentes é definida como laranja.

```js
// This code sample shows how to find and highlight the precedents 
// and direct precedents of the currently selected cell.
await Excel.run(async (context) => {
  let range = context.workbook.getActiveCell();
  // Precedents are all cells that provide data to the selected formula.
  let precedents = range.getPrecedents();
  // Direct precedents are the parent cells, or the first preceding group of cells that provide data to the selected formula.    
  let directPrecedents = range.getDirectPrecedents();

  range.load("address");
  precedents.areas.load("address");
  directPrecedents.areas.load("address");
  
  await context.sync();

  console.log(`All precedent cells of ${range.address}:`);
  
  // Use the precedents API to loop through all precedents of the active cell.
  for (let i = 0; i < precedents.areas.items.length; i++) {
    // Highlight and print out the address of all precedent cells.
    precedents.areas.items[i].format.fill.color = "Orange";
    console.log(`  ${precedents.areas.items[i].address}`);
  }

  console.log(`Direct precedent cells of ${range.address}:`);

  // Use the direct precedents API to loop through direct precedents of the active cell.
  for (let i = 0; i < directPrecedents.areas.items.length; i++) {
    // Highlight and print out the address of each direct precedent cell.
    directPrecedents.areas.items[i].format.fill.color = "Yellow";
    console.log(`  ${directPrecedents.areas.items[i].address}`);
  }
});
```

## <a name="get-the-direct-dependents-of-a-formula"></a>Obter os dependentes diretos de uma fórmula

Localize as células dependentes diretas de uma fórmula com [Range.getDirectDependents](/javascript/api/excel/excel.range#excel-excel-range-getdirectdependents-member(1)). Como `Range.getDirectPrecedents`, `Range.getDirectDependents` também retorna um `WorkbookRangeAreas` objeto. Este objeto contém os endereços de todos os dependentes diretos na guia de trabalho. Ele tem um objeto separado `RangeAreas` para cada planilha que contém pelo menos uma fórmula dependente. Para obter mais informações sobre como trabalhar com o `RangeAreas` objeto, consulte [Work with multiple ranges simultaneously in Excel add-ins](excel-add-ins-multiple-ranges.md).

A captura de tela a seguir mostra o resultado da seleção do botão **Rastrear Dependentes** na interface Excel interface do usuário. Este botão desenha uma seta de células dependentes para a célula selecionada. A célula selecionada, **D3**, tem a **célula E3** como dependente. **O E3** contém a fórmula "=C3 * D3". Ao contrário do Excel da interface do usuário, o `getDirectDependents` método não desenha setas.

![Seta rastreando células dependentes na interface Excel interface do usuário.](../images/excel-ranges-trace-dependents.png)

> [!IMPORTANT]
> O `getDirectDependents` método não recupera células dependentes entre as guias de trabalho.

O exemplo de código a seguir obtém os dependentes diretos do intervalo ativo e, em seguida, altera a cor de plano de fundo dessas células dependentes para amarelo.

```js
// This code sample shows how to find and highlight the dependents of the currently selected cell.
await Excel.run(async (context) => {
    // Direct dependents are cells that contain formulas that refer to other cells.
    let range = context.workbook.getActiveCell();
    let directDependents = range.getDirectDependents();
    range.load("address");
    directDependents.areas.load("address");
    
    await context.sync();
    console.log(`Direct dependent cells of ${range.address}:`);

    // Use the direct dependents API to loop through direct dependents of the active cell.
    for (let i = 0; i < directDependents.areas.items.length; i++) {
      // Highlight and print the address of each dependent cell.
      directDependents.areas.items[i].format.fill.color = "Yellow";
      console.log(`  ${directDependents.areas.items[i].address}`);
    }
});
```

## <a name="see-also"></a>Confira também

- [Modelo de objeto JavaScript do Excel em Suplementos do Office](excel-add-ins-core-concepts.md)
- [Trabalhar com células usando a EXCEL JavaScript](excel-add-ins-cells.md)
- [Trabalhar simultaneamente com vários intervalos em suplementos do Excel](excel-add-ins-multiple-ranges.md)
