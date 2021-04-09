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
# <a name="get-formula-precedents-using-the-excel-javascript-api"></a><span data-ttu-id="71df3-103">Obter precedentes de fórmula usando a API JavaScript do Excel</span><span class="sxs-lookup"><span data-stu-id="71df3-103">Get formula precedents using the Excel JavaScript API</span></span>

<span data-ttu-id="71df3-104">Este artigo fornece um exemplo de código que recupera precedentes de fórmula usando a API JavaScript do Excel.</span><span class="sxs-lookup"><span data-stu-id="71df3-104">This article provides a code sample that retrieves formula precedents using the Excel JavaScript API.</span></span> <span data-ttu-id="71df3-105">Para ver a lista completa de propriedades e métodos compatíveis com o `Range` objeto, consulte [Classe Excel.Range](/javascript/api/excel/excel.range).</span><span class="sxs-lookup"><span data-stu-id="71df3-105">For the complete list of properties and methods that the `Range` object supports, see [Excel.Range class](/javascript/api/excel/excel.range).</span></span>

## <a name="get-formula-precedents"></a><span data-ttu-id="71df3-106">Obter precedentes de fórmula</span><span class="sxs-lookup"><span data-stu-id="71df3-106">Get formula precedents</span></span>

<span data-ttu-id="71df3-107">Uma fórmula do Excel geralmente se refere a outras células.</span><span class="sxs-lookup"><span data-stu-id="71df3-107">An Excel formula often refers to other cells.</span></span> <span data-ttu-id="71df3-108">Quando uma célula fornece dados a uma fórmula, ela é conhecida como uma fórmula "precedente".</span><span class="sxs-lookup"><span data-stu-id="71df3-108">When a cell provides data to a formula, it is known as a formula "precedent".</span></span> <span data-ttu-id="71df3-109">Para saber mais sobre os recursos do Excel relacionados às relações entre células, consulte Exibir as relações entre [fórmulas e células.](https://support.microsoft.com/office/display-the-relationships-between-formulas-and-cells-a59bef2b-3701-46bf-8ff1-d3518771d507)</span><span class="sxs-lookup"><span data-stu-id="71df3-109">To learn more about Excel features related to relationships between cells, see [Display the relationships between formulas and cells](https://support.microsoft.com/office/display-the-relationships-between-formulas-and-cells-a59bef2b-3701-46bf-8ff1-d3518771d507).</span></span> 

<span data-ttu-id="71df3-110">Com [Range.getDirectPrecedents,](/javascript/api/excel/excel.range#getdirectprecedents--)seu complemento pode localizar células precedentes diretas de uma fórmula.</span><span class="sxs-lookup"><span data-stu-id="71df3-110">With [Range.getDirectPrecedents](/javascript/api/excel/excel.range#getdirectprecedents--), your add-in can locate a formula's direct precedent cells.</span></span> <span data-ttu-id="71df3-111">`Range.getDirectPrecedents` retorna um `WorkbookRangeAreas` objeto.</span><span class="sxs-lookup"><span data-stu-id="71df3-111">`Range.getDirectPrecedents` returns a `WorkbookRangeAreas` object.</span></span> <span data-ttu-id="71df3-112">Este objeto contém os endereços de todos os precedentes na workbook.</span><span class="sxs-lookup"><span data-stu-id="71df3-112">This object contains the addresses of all the precedents in the workbook.</span></span> <span data-ttu-id="71df3-113">Ele tem um objeto `RangeAreas` separado para cada planilha que contém pelo menos um precedente de fórmula.</span><span class="sxs-lookup"><span data-stu-id="71df3-113">It has a separate `RangeAreas` object for each worksheet containing at least one formula precedent.</span></span> <span data-ttu-id="71df3-114">Para obter mais informações sobre como trabalhar com o `RangeAreas` objeto, consulte [Work with multiple ranges simultaneously in Excel add-ins](excel-add-ins-multiple-ranges.md).</span><span class="sxs-lookup"><span data-stu-id="71df3-114">For more information on working with the `RangeAreas` object, see [Work with multiple ranges simultaneously in Excel add-ins](excel-add-ins-multiple-ranges.md).</span></span>

<span data-ttu-id="71df3-115">Na interface do usuário do Excel, o botão **Rastrear Precedentes** desenha uma seta das células precedentes para a fórmula selecionada.</span><span class="sxs-lookup"><span data-stu-id="71df3-115">In the Excel UI, the **Trace Precedents** button draws an arrow from precedent cells to the selected formula.</span></span> <span data-ttu-id="71df3-116">Ao contrário do botão da interface do usuário do Excel, `getDirectPrecedents` o método não desenha setas.</span><span class="sxs-lookup"><span data-stu-id="71df3-116">Unlike the Excel UI button, the `getDirectPrecedents` method does not draw arrows.</span></span> 

> [!IMPORTANT]
> <span data-ttu-id="71df3-117">O `getDirectPrecedents` método não pode recuperar células precedentes entre as guias de trabalho.</span><span class="sxs-lookup"><span data-stu-id="71df3-117">The `getDirectPrecedents` method can't retrieve precedent cells across workbooks.</span></span> 

<span data-ttu-id="71df3-118">O exemplo de código a seguir obtém os precedentes diretos para o intervalo ativo e, em seguida, altera a cor de plano de fundo dessas células precedentes para amarelo.</span><span class="sxs-lookup"><span data-stu-id="71df3-118">The following code sample gets the direct precedents for the active range and then changes the background color of those precedent cells to yellow.</span></span> 

> [!NOTE]
> <span data-ttu-id="71df3-119">O intervalo ativo deve conter uma fórmula que faz referência a outras células na mesma manual de trabalho para que o realçamento funcione corretamente.</span><span class="sxs-lookup"><span data-stu-id="71df3-119">The active range must contain a formula that references other cells in the same workbook for the highlighting to work properly.</span></span> 

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

## <a name="see-also"></a><span data-ttu-id="71df3-120">Confira também</span><span class="sxs-lookup"><span data-stu-id="71df3-120">See also</span></span>

- [<span data-ttu-id="71df3-121">Modelo de objeto JavaScript do Excel em Suplementos do Office</span><span class="sxs-lookup"><span data-stu-id="71df3-121">Excel JavaScript object model in Office Add-ins</span></span>](excel-add-ins-core-concepts.md)
- [<span data-ttu-id="71df3-122">Trabalhar com células usando a API JavaScript do Excel</span><span class="sxs-lookup"><span data-stu-id="71df3-122">Work with cells using the Excel JavaScript API</span></span>](excel-add-ins-cells.md)
- [<span data-ttu-id="71df3-123">Trabalhar simultaneamente com vários intervalos em suplementos do Excel</span><span class="sxs-lookup"><span data-stu-id="71df3-123">Work with multiple ranges simultaneously in Excel add-ins</span></span>](excel-add-ins-multiple-ranges.md)
