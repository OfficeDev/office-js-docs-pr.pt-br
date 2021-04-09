---
title: Trabalhar com precedentes de fórmula usando a API JavaScript do Excel
description: Saiba como usar a API JavaScript do Excel para recuperar precedentes de fórmula.
ms.date: 04/02/2021
ms.prod: excel
localization_priority: Normal
ms.openlocfilehash: 0d21ae411615a22873a0f4dda185984f6191ac8e
ms.sourcegitcommit: 54fef33bfc7d18a35b3159310bbd8b1c8312f845
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 04/09/2021
ms.locfileid: "51652766"
---
# <a name="get-formula-precedents-using-the-excel-javascript-api"></a>Obter precedentes de fórmula usando a API JavaScript do Excel

Este artigo fornece um exemplo de código que recupera precedentes de fórmula usando a API JavaScript do Excel. Para ver a lista completa de propriedades e métodos compatíveis com o `Range` objeto, consulte [Classe Excel.Range](/javascript/api/excel/excel.range).

## <a name="get-formula-precedents"></a>Obter precedentes de fórmula

Uma fórmula do Excel geralmente se refere a outras células. Quando uma célula fornece dados a uma fórmula, ela é conhecida como uma fórmula "precedente". Para saber mais sobre os recursos do Excel relacionados às relações entre células, consulte Exibir as relações entre [fórmulas e células.](https://support.microsoft.com/office/display-the-relationships-between-formulas-and-cells-a59bef2b-3701-46bf-8ff1-d3518771d507) 

Com [Range.getDirectPrecedents,](/javascript/api/excel/excel.range#getdirectprecedents--)seu complemento pode localizar células precedentes diretas de uma fórmula. `Range.getDirectPrecedents` retorna um `WorkbookRangeAreas` objeto. Este objeto contém os endereços de todos os precedentes na workbook. Ele tem um objeto `RangeAreas` separado para cada planilha que contém pelo menos um precedente de fórmula. Para obter mais informações sobre como trabalhar com o `RangeAreas` objeto, consulte [Work with multiple ranges simultaneously in Excel add-ins](excel-add-ins-multiple-ranges.md).

Na interface do usuário do Excel, o botão **Rastrear Precedentes** desenha uma seta das células precedentes para a fórmula selecionada. Ao contrário do botão da interface do usuário do Excel, `getDirectPrecedents` o método não desenha setas. 

> [!IMPORTANT]
> O `getDirectPrecedents` método não pode recuperar células precedentes entre as guias de trabalho. 

O exemplo de código a seguir obtém os precedentes diretos para o intervalo ativo e, em seguida, altera a cor de plano de fundo dessas células precedentes para amarelo. 

> [!NOTE]
> O intervalo ativo deve conter uma fórmula que faz referência a outras células na mesma manual de trabalho para que o realçamento funcione corretamente. 

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
        })
        .then(context.sync);
}).catch(errorHandlerFunction);
```

## <a name="see-also"></a>Confira também

- [Modelo de objeto JavaScript do Excel em Suplementos do Office](excel-add-ins-core-concepts.md)
- [Trabalhar com células usando a API JavaScript do Excel](excel-add-ins-cells.md)
- [Trabalhar simultaneamente com vários intervalos em suplementos do Excel](excel-add-ins-multiple-ranges.md)
