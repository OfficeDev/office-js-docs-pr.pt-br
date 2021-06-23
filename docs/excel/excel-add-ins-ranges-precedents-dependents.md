---
title: Trabalhar com precedentes de fórmula e dependentes usando Excel API JavaScript
description: Saiba como usar a API JavaScript Excel para recuperar precedentes e dependentes da fórmula.
ms.date: 06/03/2021
ms.prod: excel
localization_priority: Normal
ms.openlocfilehash: 6021e383f02ca0de15210638b991dfe8b109ab63
ms.sourcegitcommit: ee9e92a968e4ad23f1e371f00d4888e4203ab772
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 06/23/2021
ms.locfileid: "53075793"
---
# <a name="get-formula-precedents-and-dependents-using-the-excel-javascript-api"></a><span data-ttu-id="6bb2d-103">Obter precedentes de fórmula e dependentes usando a API JavaScript Excel javascript</span><span class="sxs-lookup"><span data-stu-id="6bb2d-103">Get formula precedents and dependents using the Excel JavaScript API</span></span>

<span data-ttu-id="6bb2d-104">Excel fórmulas geralmente se referem a outras células.</span><span class="sxs-lookup"><span data-stu-id="6bb2d-104">Excel formulas often refer to other cells.</span></span> <span data-ttu-id="6bb2d-105">Essas referências entre células são conhecidas como "precedentes" e "dependentes".</span><span class="sxs-lookup"><span data-stu-id="6bb2d-105">These cross-cell references are known as "precedents" and "dependents".</span></span> <span data-ttu-id="6bb2d-106">Um precedente é uma célula que fornece dados a uma fórmula.</span><span class="sxs-lookup"><span data-stu-id="6bb2d-106">A precedent is a cell that provides data to a formula.</span></span> <span data-ttu-id="6bb2d-107">Um dependente é uma célula que contém uma fórmula que se refere a outras células.</span><span class="sxs-lookup"><span data-stu-id="6bb2d-107">A dependent is a cell that contains a formula that refers to other cells.</span></span> <span data-ttu-id="6bb2d-108">Para saber mais sobre os Excel relacionados às relações entre células, consulte Exibir as relações entre [fórmulas e células.](https://support.microsoft.com/office/display-the-relationships-between-formulas-and-cells-a59bef2b-3701-46bf-8ff1-d3518771d507)</span><span class="sxs-lookup"><span data-stu-id="6bb2d-108">To learn more about Excel features related to relationships between cells, see [Display the relationships between formulas and cells](https://support.microsoft.com/office/display-the-relationships-between-formulas-and-cells-a59bef2b-3701-46bf-8ff1-d3518771d507).</span></span>

<span data-ttu-id="6bb2d-109">Uma célula pode ter uma célula precedente, e essa célula precedente pode ter suas próprias células precedentes.</span><span class="sxs-lookup"><span data-stu-id="6bb2d-109">A cell may have a precedent cell, and that precedent cell may have its own precedent cells.</span></span> <span data-ttu-id="6bb2d-110">Um "precedente direto" é o primeiro grupo de células anterior nesta sequência, semelhante ao conceito de pais em uma relação pai-filho.</span><span class="sxs-lookup"><span data-stu-id="6bb2d-110">A "direct precedent" is the first preceding group of cells in this sequence, similar to the concept of parents in a parent-child relationship.</span></span> <span data-ttu-id="6bb2d-111">Um "dependente direto" é o primeiro grupo dependente de células em uma sequência, semelhante a filhos em uma relação pai-filho.</span><span class="sxs-lookup"><span data-stu-id="6bb2d-111">A "direct dependent" is the first dependent group of cells in a sequence, similar to children in a parent-child relationship.</span></span> <span data-ttu-id="6bb2d-112">Células que se referem a outras células em uma workbook, mas cuja relação não é uma relação pai-filho, não são dependentes diretos ou precedentes diretos.</span><span class="sxs-lookup"><span data-stu-id="6bb2d-112">Cells that refer to other cells in a workbook, but whose relationship is not a parent-child relationship, are not direct dependents or direct precedents.</span></span>

<span data-ttu-id="6bb2d-113">Este artigo fornece exemplos de código que recuperam precedentes diretos e dependentes diretos de fórmulas usando Excel API JavaScript.</span><span class="sxs-lookup"><span data-stu-id="6bb2d-113">This article provides code samples that retrieve direct precedents and direct dependents of formulas using the Excel JavaScript API.</span></span> <span data-ttu-id="6bb2d-114">Para ver a lista completa de propriedades e métodos que o objeto oferece suporte, consulte `Range` [Range Object (API JavaScript para Excel)](/javascript/api/excel/excel.range).</span><span class="sxs-lookup"><span data-stu-id="6bb2d-114">For the complete list of properties and methods that the `Range` object supports, see [Range Object (JavaScript API for Excel)](/javascript/api/excel/excel.range).</span></span>

## <a name="get-the-direct-precedents-of-a-formula"></a><span data-ttu-id="6bb2d-115">Obter os precedentes diretos de uma fórmula</span><span class="sxs-lookup"><span data-stu-id="6bb2d-115">Get the direct precedents of a formula</span></span>

<span data-ttu-id="6bb2d-116">Localize as células precedentes diretas de uma fórmula [com Range.getDirectPrecedents](/javascript/api/excel/excel.range#getdirectprecedents--).</span><span class="sxs-lookup"><span data-stu-id="6bb2d-116">Locate a formula's direct precedent cells with [Range.getDirectPrecedents](/javascript/api/excel/excel.range#getdirectprecedents--).</span></span> <span data-ttu-id="6bb2d-117">`Range.getDirectPrecedents` retorna um `WorkbookRangeAreas` objeto.</span><span class="sxs-lookup"><span data-stu-id="6bb2d-117">`Range.getDirectPrecedents` returns a `WorkbookRangeAreas` object.</span></span> <span data-ttu-id="6bb2d-118">Este objeto contém os endereços de todos os precedentes diretos na guia de trabalho.</span><span class="sxs-lookup"><span data-stu-id="6bb2d-118">This object contains the addresses of all the direct precedents in the workbook.</span></span> <span data-ttu-id="6bb2d-119">Ele tem um objeto `RangeAreas` separado para cada planilha que contém pelo menos um precedente de fórmula.</span><span class="sxs-lookup"><span data-stu-id="6bb2d-119">It has a separate `RangeAreas` object for each worksheet containing at least one formula precedent.</span></span> <span data-ttu-id="6bb2d-120">Para obter mais informações sobre como trabalhar com o objeto, consulte `RangeAreas` Work with multiple [ranges simultaneously in Excel add-ins](excel-add-ins-multiple-ranges.md).</span><span class="sxs-lookup"><span data-stu-id="6bb2d-120">For more information on working with the `RangeAreas` object, see [Work with multiple ranges simultaneously in Excel add-ins](excel-add-ins-multiple-ranges.md).</span></span>

<span data-ttu-id="6bb2d-121">A captura de tela a seguir mostra o resultado da seleção do botão **Rastrear Precedentes** na interface Excel interface do usuário.</span><span class="sxs-lookup"><span data-stu-id="6bb2d-121">The following screenshot shows the result of selecting the **Trace Precedents** button in the Excel UI.</span></span> <span data-ttu-id="6bb2d-122">Este botão desenha uma seta de células precedentes para a célula selecionada.</span><span class="sxs-lookup"><span data-stu-id="6bb2d-122">This button draws an arrow from precedent cells to the selected cell.</span></span> <span data-ttu-id="6bb2d-123">A célula selecionada, **E3**, contém a fórmula "=C3 \* D3", **portanto, C3** e **D3** são células precedentes.</span><span class="sxs-lookup"><span data-stu-id="6bb2d-123">The selected cell, **E3**, contains the formula "=C3 \* D3", so both **C3** and **D3** are precedent cells.</span></span> <span data-ttu-id="6bb2d-124">Ao contrário do Excel da interface do usuário, o `getDirectPrecedents` método não desenha setas.</span><span class="sxs-lookup"><span data-stu-id="6bb2d-124">Unlike the Excel UI button, the `getDirectPrecedents` method does not draw arrows.</span></span>

![Seta rastreando células precedentes na interface Excel interface do usuário.](../images/excel-ranges-trace-precedents.png)

> [!IMPORTANT]
> <span data-ttu-id="6bb2d-126">O `getDirectPrecedents` método não pode recuperar células precedentes entre as guias de trabalho.</span><span class="sxs-lookup"><span data-stu-id="6bb2d-126">The `getDirectPrecedents` method can't retrieve precedent cells across workbooks.</span></span>

<span data-ttu-id="6bb2d-127">O exemplo de código a seguir obtém os precedentes diretos para o intervalo ativo e, em seguida, altera a cor de plano de fundo dessas células precedentes para amarelo.</span><span class="sxs-lookup"><span data-stu-id="6bb2d-127">The following code sample gets the direct precedents for the active range and then changes the background color of those precedent cells to yellow.</span></span>

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

## <a name="get-the-direct-dependents-of-a-formula-preview"></a><span data-ttu-id="6bb2d-128">Obter os dependentes diretos de uma fórmula (visualização)</span><span class="sxs-lookup"><span data-stu-id="6bb2d-128">Get the direct dependents of a formula (preview)</span></span>

> [!NOTE]
> <span data-ttu-id="6bb2d-129">No `Range.getDirectDependents` momento, o método só está disponível na visualização pública.</span><span class="sxs-lookup"><span data-stu-id="6bb2d-129">The `Range.getDirectDependents` method is currently only available in public preview.</span></span> [!INCLUDE [Information about using preview APIs](../includes/using-excel-preview-apis.md)]
> 

<span data-ttu-id="6bb2d-130">Localize as células dependentes diretas de uma fórmula [com Range.getDirectDependents](/javascript/api/excel/excel.range#getDirectDependents__).</span><span class="sxs-lookup"><span data-stu-id="6bb2d-130">Locate a formula's direct dependent cells with [Range.getDirectDependents](/javascript/api/excel/excel.range#getDirectDependents__).</span></span> <span data-ttu-id="6bb2d-131">Como `Range.getDirectPrecedents` , também retorna um `Range.getDirectDependents` `WorkbookRangeAreas` objeto.</span><span class="sxs-lookup"><span data-stu-id="6bb2d-131">Like `Range.getDirectPrecedents`, `Range.getDirectDependents` also returns a `WorkbookRangeAreas` object.</span></span> <span data-ttu-id="6bb2d-132">Este objeto contém os endereços de todos os dependentes diretos na guia de trabalho.</span><span class="sxs-lookup"><span data-stu-id="6bb2d-132">This object contains the addresses of all the direct dependents in the workbook.</span></span> <span data-ttu-id="6bb2d-133">Ele tem um objeto `RangeAreas` separado para cada planilha que contém pelo menos uma fórmula dependente.</span><span class="sxs-lookup"><span data-stu-id="6bb2d-133">It has a separate `RangeAreas` object for each worksheet containing at least one formula dependent.</span></span> <span data-ttu-id="6bb2d-134">Para obter mais informações sobre como trabalhar com o objeto, consulte `RangeAreas` Work with multiple [ranges simultaneously in Excel add-ins](excel-add-ins-multiple-ranges.md).</span><span class="sxs-lookup"><span data-stu-id="6bb2d-134">For more information on working with the `RangeAreas` object, see [Work with multiple ranges simultaneously in Excel add-ins](excel-add-ins-multiple-ranges.md).</span></span>

<span data-ttu-id="6bb2d-135">A captura de tela a seguir mostra o resultado da seleção do botão **Rastrear Dependentes** na interface Excel interface do usuário.</span><span class="sxs-lookup"><span data-stu-id="6bb2d-135">The following screenshot shows the result of selecting the **Trace Dependents** button in the Excel UI.</span></span> <span data-ttu-id="6bb2d-136">Este botão desenha uma seta de células dependentes para a célula selecionada.</span><span class="sxs-lookup"><span data-stu-id="6bb2d-136">This button draws an arrow from dependent cells to the selected cell.</span></span> <span data-ttu-id="6bb2d-137">A célula selecionada, **D3**, tem a **célula E3** como dependente.</span><span class="sxs-lookup"><span data-stu-id="6bb2d-137">The selected cell, **D3**, has cell **E3** as a dependent.</span></span> <span data-ttu-id="6bb2d-138">**O E3** contém a fórmula "=C3 \* D3".</span><span class="sxs-lookup"><span data-stu-id="6bb2d-138">**E3** contains the formula "=C3 \* D3".</span></span> <span data-ttu-id="6bb2d-139">Ao contrário do Excel da interface do usuário, o `getDirectDependents` método não desenha setas.</span><span class="sxs-lookup"><span data-stu-id="6bb2d-139">Unlike the Excel UI button, the `getDirectDependents` method does not draw arrows.</span></span>

![Células dependentes de rastreamento de seta na interface Excel interface do usuário.](../images/excel-ranges-trace-dependents.png)

> [!IMPORTANT]
> <span data-ttu-id="6bb2d-141">O `getDirectDependents` método não pode recuperar células dependentes entre as guias de trabalho.</span><span class="sxs-lookup"><span data-stu-id="6bb2d-141">The `getDirectDependents` method can't retrieve dependent cells across workbooks.</span></span>

<span data-ttu-id="6bb2d-142">O exemplo de código a seguir obtém os dependentes diretos do intervalo ativo e, em seguida, altera a cor de plano de fundo dessas células dependentes para amarelo.</span><span class="sxs-lookup"><span data-stu-id="6bb2d-142">The following code sample gets the direct dependents for the active range and then changes the background color of those dependent cells to yellow.</span></span>

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

## <a name="see-also"></a><span data-ttu-id="6bb2d-143">Confira também</span><span class="sxs-lookup"><span data-stu-id="6bb2d-143">See also</span></span>

- [<span data-ttu-id="6bb2d-144">Modelo de objeto JavaScript do Excel em Suplementos do Office</span><span class="sxs-lookup"><span data-stu-id="6bb2d-144">Excel JavaScript object model in Office Add-ins</span></span>](excel-add-ins-core-concepts.md)
- [<span data-ttu-id="6bb2d-145">Trabalhar com células usando a EXCEL JavaScript</span><span class="sxs-lookup"><span data-stu-id="6bb2d-145">Work with cells using the Excel JavaScript API</span></span>](excel-add-ins-cells.md)
- [<span data-ttu-id="6bb2d-146">Trabalhar simultaneamente com vários intervalos em suplementos do Excel</span><span class="sxs-lookup"><span data-stu-id="6bb2d-146">Work with multiple ranges simultaneously in Excel add-ins</span></span>](excel-add-ins-multiple-ranges.md)
