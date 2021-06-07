---
title: Trabalhar com precedentes de fórmula e dependentes usando Excel API JavaScript
description: Saiba como usar a API JavaScript Excel para recuperar precedentes e dependentes da fórmula.
ms.date: 06/03/2021
ms.prod: excel
localization_priority: Normal
ms.openlocfilehash: 78fa4fb070ede85d139425a9d59ba1224785a605
ms.sourcegitcommit: 17b5a076375bc5dc3f91d3602daeb7535d67745d
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 06/06/2021
ms.locfileid: "52783516"
---
# <a name="get-formula-precedents-and-dependents-using-the-excel-javascript-api"></a>Obter precedentes de fórmula e dependentes usando a API JavaScript Excel javascript

Excel fórmulas geralmente se referem a outras células. Essas referências entre células são conhecidas como "precedentes" e "dependentes". Um precedente é uma célula que fornece dados a uma fórmula. Um dependente é uma célula que contém uma fórmula que se refere a outras células. Para saber mais sobre os Excel relacionados às relações entre células, consulte Exibir as relações entre [fórmulas e células.](https://support.microsoft.com/office/display-the-relationships-between-formulas-and-cells-a59bef2b-3701-46bf-8ff1-d3518771d507)

Uma célula pode ter uma célula precedente, e essa célula precedente pode ter suas próprias células precedentes. Um "precedente direto" é o primeiro grupo de células anterior nesta sequência, semelhante ao conceito de pais em uma relação pai-filho. Um "dependente direto" é o primeiro grupo dependente de células em uma sequência, semelhante a filhos em uma relação pai-filho. Células que se referem a outras células em uma workbook, mas cuja relação não é uma relação pai-filho, não são dependentes diretos ou precedentes diretos.

Este artigo fornece exemplos de código que recuperam precedentes diretos e dependentes diretos de fórmulas usando Excel API JavaScript. Para ver a lista completa de propriedades e métodos que o objeto oferece suporte, consulte `Range` [Range Object (API JavaScript para Excel)](/javascript/api/excel/excel.range).

## <a name="get-the-direct-precedents-of-a-formula"></a>Obter os precedentes diretos de uma fórmula

Localize as células precedentes diretas de uma fórmula [com Range.getDirectPrecedents](/javascript/api/excel/excel.range#getdirectprecedents--). `Range.getDirectPrecedents` retorna um `WorkbookRangeAreas` objeto. Este objeto contém os endereços de todos os precedentes diretos na guia de trabalho. Ele tem um objeto `RangeAreas` separado para cada planilha que contém pelo menos um precedente de fórmula. Para obter mais informações sobre como trabalhar com o objeto, consulte `RangeAreas` Work with multiple [ranges simultaneously in Excel add-ins](excel-add-ins-multiple-ranges.md).

A captura de tela a seguir mostra o resultado da seleção do botão **Rastrear Precedentes** na interface Excel interface do usuário. Este botão desenha uma seta de células precedentes para a célula selecionada. A célula selecionada, **E3**, contém a fórmula "=C3 * D3", **portanto, C3** e **D3** são células precedentes. Ao contrário do Excel da interface do usuário, o `getDirectPrecedents` método não desenha setas.

![Células precedentes de rastreamento de seta na interface do usuário Excel de seta](../images/excel-ranges-trace-precedents.png)

> [!IMPORTANT]
> O `getDirectPrecedents` método não pode recuperar células precedentes entre as guias de trabalho.

O exemplo de código a seguir obtém os precedentes diretos para o intervalo ativo e, em seguida, altera a cor de plano de fundo dessas células precedentes para amarelo.

```js
Excel.run(function (context) {
    // Precedents are cells that provide data to the selected formula.
    var range = context.workbook.getActiveCell();
    var directPrecedents = range.getDirectPrecedents();
    range.load("address");
    directPrecedents.areas.load("address");
    
    return context.sync()
        .then(function () {
            console.log(`Direct precedent cells of ${range.address}:`);

            // Use the direct precedents API to loop through precedents of the active cell.
            for (var i = 0; i < directPrecedents.areas.items.length; i++) {
              // Highlight and print out the address of each precedent cell.
              directPrecedents.areas.items[i].format.fill.color = "Yellow";
              console.log(`  ${directPrecedents.areas.items[i].address}`);
            }
        });
}).catch(errorHandlerFunction);
```

## <a name="get-the-direct-dependents-of-a-formula-preview"></a>Obter os dependentes diretos de uma fórmula (visualização)

> [!NOTE]
> No `Range.getDirectDependents` momento, o método só está disponível na visualização pública. [!INCLUDE [Information about using preview APIs](../includes/using-excel-preview-apis.md)]
> 

Localize as células dependentes diretas de uma fórmula [com Range.getDirectDependents](/javascript/api/excel/excel.range#getDirectDependents__). Como `Range.getDirectPrecedents` , também retorna um `Range.getDirectDependents` `WorkbookRangeAreas` objeto. Este objeto contém os endereços de todos os dependentes diretos na guia de trabalho. Ele tem um objeto `RangeAreas` separado para cada planilha que contém pelo menos uma fórmula dependente. Para obter mais informações sobre como trabalhar com o objeto, consulte `RangeAreas` Work with multiple [ranges simultaneously in Excel add-ins](excel-add-ins-multiple-ranges.md).

A captura de tela a seguir mostra o resultado da seleção do botão **Rastrear Dependentes** na interface Excel interface do usuário. Este botão desenha uma seta de células dependentes para a célula selecionada. A célula selecionada, **D3**, tem a **célula E3** como dependente. **O E3** contém a fórmula "=C3 * D3". Ao contrário do Excel da interface do usuário, o `getDirectDependents` método não desenha setas.

![Células dependentes de rastreamento de seta na interface Excel interface do usuário](../images/excel-ranges-trace-dependents.png)

> [!IMPORTANT]
> O `getDirectDependents` método não pode recuperar células dependentes entre as guias de trabalho.

O exemplo de código a seguir obtém os dependentes diretos do intervalo ativo e, em seguida, altera a cor de plano de fundo dessas células dependentes para amarelo.

```js
Excel.run(function (context) {
    // Direct dependents are cells that contain formulas that refer to other cells.
    var range = context.workbook.getActiveCell();
    var directDependents = range.getDirectDependents();
    range.load("address");
    directDependents.areas.load("address");
    
    return context.sync()
        .then(function () {
            console.log(`Direct dependent cells of ${range.address}:`);
    
            // Use the direct dependents API to loop through direct dependents of the active cell.
            for (var i = 0; i < directDependents.areas.items.length; i++) {
              // Highlight and print the address of each dependent cell.
              directDependents.areas.items[i].format.fill.color = "Yellow";
              console.log(`  ${directDependents.areas.items[i].address}`);
            }
        });
}).catch(errorHandlerFunction);
```

## <a name="see-also"></a>Confira também

- [Modelo de objeto JavaScript do Excel em Suplementos do Office](excel-add-ins-core-concepts.md)
- [Trabalhar com células usando a EXCEL JavaScript](excel-add-ins-cells.md)
- [Trabalhar simultaneamente com vários intervalos em suplementos do Excel](excel-add-ins-multiple-ranges.md)
