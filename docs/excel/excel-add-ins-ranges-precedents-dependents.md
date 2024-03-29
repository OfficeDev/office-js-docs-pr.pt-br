---
title: Trabalhar com precedentes de fórmula e dependentes usando a API Excel JavaScript
description: Saiba como usar a API JavaScript Excel para recuperar precedentes e dependentes de fórmula.
ms.date: 05/19/2022
ms.prod: excel
ms.localizationpriority: medium
ms.openlocfilehash: ca432b7eb6825781960e995af2ed2193c7caa5e2
ms.sourcegitcommit: 4ca3334f3cefa34e6b391eb92a429a308229fe89
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 05/21/2022
ms.locfileid: "65628093"
---
# <a name="get-formula-precedents-and-dependents-using-the-excel-javascript-api"></a>Obter precedentes de fórmula e dependentes usando a API Excel JavaScript

Excel fórmulas geralmente se referem a outras células. Essas referências entre células são conhecidas como "precedentes" e "dependentes". Um precedente é uma célula que fornece dados para uma fórmula. Um dependente é uma célula que contém uma fórmula que se refere a outras células. Para saber mais sobre Excel relacionados a relações entre células, consulte Exibir as relações entre [fórmulas e células](https://support.microsoft.com/office/a59bef2b-3701-46bf-8ff1-d3518771d507).

Uma célula precedente pode ter suas próprias células precedentes. Cada célula precedente nesta cadeia de precedentes ainda é um precedente da célula original. A mesma relação existe para dependentes. Qualquer célula afetada por outra célula é dependente dessa célula. Um "precedente direto" é o primeiro grupo anterior de células nesta sequência, semelhante ao conceito de pais em uma relação pai-filho. Um "dependente direto" é o primeiro grupo dependente de células em uma sequência, semelhante aos filhos em uma relação pai-filho.

Este artigo fornece exemplos de código que recuperam precedentes e dependentes de fórmulas usando Excel API JavaScript. Para obter a lista completa de propriedades e métodos compatíveis `Range` com o objeto, consulte [Objeto Range (API JavaScript para Excel)](/javascript/api/excel/excel.range).

## <a name="get-the-precedents-of-a-formula"></a>Obter os precedentes de uma fórmula

Localize as células precedentes de uma fórmula [com Range.getPrecedents](/javascript/api/excel/excel.range#excel-excel-range-getprecedents-member(1)). `Range.getPrecedents` retorna um `WorkbookRangeAreas` objeto. Esse objeto contém os endereços de todos os precedentes na pasta de trabalho. Ele tem um objeto separado `RangeAreas` para cada planilha que contém pelo menos uma fórmula precedente. Para saber mais sobre o objeto`RangeAreas`, consulte [Trabalhar com vários intervalos simultaneamente Excel suplementos](excel-add-ins-multiple-ranges.md).

Para localizar apenas as células precedentes diretas de uma fórmula, use [Range.getDirectPrecedents](/javascript/api/excel/excel.range#excel-excel-range-getdirectprecedents-member(1)). `Range.getDirectPrecedents` funciona como `Range.getPrecedents` e retorna um `WorkbookRangeAreas` objeto que contém os endereços de precedentes diretos.

A captura de tela a seguir mostra o resultado da seleção do botão **Rastrear Precedentes** na Excel interface do usuário. Esse botão desenha uma seta de células precedentes para a célula selecionada. A célula selecionada, **E3**, contém a fórmula "=C3 * D3", portanto, **C3** e **D3** são células precedentes. Ao contrário do Excel da interface do usuário, `getPrecedents` os métodos `getDirectPrecedents` e os métodos não desenham setas.

![Seta rastreando células precedentes na Excel interface do usuário.](../images/excel-ranges-trace-precedents.png)

> [!IMPORTANT]
> Os `getPrecedents` métodos `getDirectPrecedents` e os métodos não recuperam células precedentes entre pastas de trabalho.

O exemplo de código a seguir mostra como trabalhar com os `Range.getPrecedents` métodos `Range.getDirectPrecedents` e os métodos. O exemplo obtém os precedentes para o intervalo ativo e, em seguida, altera a cor da tela de fundo dessas células precedentes. A cor da tela de fundo das células precedentes diretas é definida como amarela e a cor da tela de fundo das outras células precedentes é definida como laranja.

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

## <a name="get-the-dependents-of-a-formula"></a>Obter os dependentes de uma fórmula

Localize as células dependentes de uma fórmula [com Range.getDependents](/javascript/api/excel/excel.range#excel-excel-range-getdependents-member(1)). Like `Range.getPrecedents`, `Range.getDependents` também retorna um `WorkbookRangeAreas` objeto. Esse objeto contém os endereços de todos os dependentes na pasta de trabalho. Ele tem um objeto `RangeAreas` separado para cada planilha que contém pelo menos uma fórmula dependente. Para obter mais informações sobre como trabalhar com o `RangeAreas` objeto, consulte Trabalhar com vários [intervalos simultaneamente Excel suplementos](excel-add-ins-multiple-ranges.md).

Para localizar apenas as células dependentes diretas de uma fórmula, use [Range.getDirectDependents](/javascript/api/excel/excel.range#excel-excel-range-getdirectdependents-member(1)). `Range.getDirectDependents` funciona como `Range.getDependents` e retorna um `WorkbookRangeAreas` objeto que contém os endereços de dependentes diretos.

A captura de tela a seguir mostra o resultado da seleção do botão Rastrear **Dependentes** na Excel interface do usuário. Esse botão desenha uma seta da célula selecionada para células dependentes. A célula selecionada, **D3**, tem a **célula E3** como dependente. **E3** contém a fórmula "=C3 * D3". Ao contrário do Excel da interface do usuário, `getDependents` os métodos `getDirectDependents` e os métodos não desenham setas.

![Células dependentes de rastreamento de seta Excel interface do usuário.](../images/excel-ranges-trace-dependents.png)

> [!IMPORTANT]
> Os `getDependents` métodos e `getDirectDependents` os métodos não recuperam células dependentes em pastas de trabalho.

O exemplo de código a seguir obtém os dependentes diretos do intervalo ativo e, em seguida, altera a cor da tela de fundo dessas células dependentes para amarelo.

O exemplo de código a seguir mostra como trabalhar com os `Range.getDependents` métodos `Range.getDirectDependents` e os métodos. O exemplo obtém os dependentes do intervalo ativo e altera a cor da tela de fundo dessas células dependentes. A cor da tela de fundo das células dependentes diretas é definida como amarela e a cor da tela de fundo das outras células dependentes é definida como laranja.

```js
// This code sample shows how to find and highlight the dependents 
// and direct dependents of the currently selected cell.
await Excel.run(async (context) => {
    let range = context.workbook.getActiveCell();
    // Dependents are all cells that contain formulas that refer to other cells.
    let dependents = range.getDependents();  
    // Direct dependents are the child cells, or the first succeeding group of cells in a sequence of cells that refer to other cells.
    let directDependents = range.getDirectDependents();

    range.load("address");
    dependents.areas.load("address");    
    directDependents.areas.load("address");
    
    await context.sync();

    console.log(`All dependent cells of ${range.address}:`);
    
    // Use the dependents API to loop through all dependents of the active cell.
    for (let i = 0; i < dependents.areas.items.length; i++) {
      // Highlight and print out the addresses of all dependent cells.
      dependents.areas.items[i].format.fill.color = "Orange";
      console.log(`  ${dependents.areas.items[i].address}`);
    }

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
