---
title: Manipular matrizes dinâmicas e vazamento de intervalo usando a API JavaScript do Excel
description: Saiba como lidar com o vazamento de matrizes dinâmicas e intervalos com a API JavaScript do Excel.
ms.date: 04/02/2021
ms.prod: excel
localization_priority: Normal
ms.openlocfilehash: c224fc336791440911519a6d24aee6c208d90c9e
ms.sourcegitcommit: 54fef33bfc7d18a35b3159310bbd8b1c8312f845
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 04/09/2021
ms.locfileid: "51652774"
---
# <a name="handle-dynamic-arrays-and-spilling-using-the-excel-javascript-api"></a><span data-ttu-id="010fb-103">Manipular matrizes dinâmicas e vazamento usando a API JavaScript do Excel</span><span class="sxs-lookup"><span data-stu-id="010fb-103">Handle dynamic arrays and spilling using the Excel JavaScript API</span></span>

<span data-ttu-id="010fb-104">Este artigo fornece um exemplo de código que lida com matrizes dinâmicas e vazamento de intervalo usando a API JavaScript do Excel.</span><span class="sxs-lookup"><span data-stu-id="010fb-104">This article provides a code sample that handles dynamic arrays and range spilling using the Excel JavaScript API.</span></span> <span data-ttu-id="010fb-105">Para ver a lista completa de propriedades e métodos compatíveis com o `Range` objeto, consulte [Classe Excel.Range](/javascript/api/excel/excel.range).</span><span class="sxs-lookup"><span data-stu-id="010fb-105">For the complete list of properties and methods that the `Range` object supports, see [Excel.Range class](/javascript/api/excel/excel.range).</span></span>

## <a name="dynamic-arrays"></a><span data-ttu-id="010fb-106">Matrizes dinâmicas</span><span class="sxs-lookup"><span data-stu-id="010fb-106">Dynamic arrays</span></span>

<span data-ttu-id="010fb-107">Algumas fórmulas do Excel [retornam matrizes dinâmicas.](https://support.microsoft.com/office/dynamic-array-formulas-and-spilled-array-behavior-205c6b06-03ba-4151-89a1-87a7eb36e531)</span><span class="sxs-lookup"><span data-stu-id="010fb-107">Some Excel formulas return [Dynamic arrays](https://support.microsoft.com/office/dynamic-array-formulas-and-spilled-array-behavior-205c6b06-03ba-4151-89a1-87a7eb36e531).</span></span> <span data-ttu-id="010fb-108">Eles preenchem os valores de várias células fora da célula original da fórmula.</span><span class="sxs-lookup"><span data-stu-id="010fb-108">These fill the values of multiple cells outside of the formula's original cell.</span></span> <span data-ttu-id="010fb-109">Esse estouro de valor é chamado de "vazamento".</span><span class="sxs-lookup"><span data-stu-id="010fb-109">This value overflow is referred to as a "spill".</span></span> <span data-ttu-id="010fb-110">Seu complemento pode encontrar o intervalo usado para um vazamento com o [método Range.getSpillingToRange.](/javascript/api/excel/excel.range#getspillingtorange--)</span><span class="sxs-lookup"><span data-stu-id="010fb-110">Your add-in can find the range used for a spill with the [Range.getSpillingToRange](/javascript/api/excel/excel.range#getspillingtorange--) method.</span></span> <span data-ttu-id="010fb-111">Há também uma [versão \*OrNullObject](..//develop/application-specific-api-model.md#ornullobject-methods-and-properties), `Range.getSpillingToRangeOrNullObject` .</span><span class="sxs-lookup"><span data-stu-id="010fb-111">There is also a [\*OrNullObject version](..//develop/application-specific-api-model.md#ornullobject-methods-and-properties), `Range.getSpillingToRangeOrNullObject`.</span></span>

<span data-ttu-id="010fb-112">O exemplo a seguir mostra uma fórmula básica que copia o conteúdo de um intervalo em uma célula, que se espalha em células vizinhas.</span><span class="sxs-lookup"><span data-stu-id="010fb-112">The following sample shows a basic formula that copies the contents of a range into a cell, which spills into neighboring cells.</span></span> <span data-ttu-id="010fb-113">Em seguida, o complemento registra o intervalo que contém o vazamento.</span><span class="sxs-lookup"><span data-stu-id="010fb-113">The add-in then logs the range that contains the spill.</span></span>

```js
Excel.run(function (context) {
    var sheet = context.workbook.worksheets.getItem("Sample");

    // Set G4 to a formula that returns a dynamic array.
    var targetCell = sheet.getRange("G4");
    targetCell.formulas = [["=A4:D4"]];

    // Get the address of the cells that the dynamic array spilled into.
    var spillRange = targetCell.getSpillingToRange();
    spillRange.load("address");

    // Sync and log the spilled-to range.
    return context.sync().then(function () {
        // This will log the range as "G4:J4".
        console.log(`Copying the table headers spilled into ${spillRange.address}.`);
    });
}).catch(errorHandlerFunction);
```

## <a name="range-spilling"></a><span data-ttu-id="010fb-114">Vazamento de intervalo</span><span class="sxs-lookup"><span data-stu-id="010fb-114">Range spilling</span></span>

<span data-ttu-id="010fb-115">Encontre a célula responsável pelo vazamento em uma determinada célula usando o [método Range.getSpillParent.](/javascript/api/excel/excel.range#getspillparent--)</span><span class="sxs-lookup"><span data-stu-id="010fb-115">Find the cell responsible for spilling into a given cell by using the [Range.getSpillParent](/javascript/api/excel/excel.range#getspillparent--) method.</span></span> <span data-ttu-id="010fb-116">Observe que `getSpillParent` só funciona quando o objeto range é uma única célula.</span><span class="sxs-lookup"><span data-stu-id="010fb-116">Note that `getSpillParent` only works when the range object is a single cell.</span></span> <span data-ttu-id="010fb-117">Chamar em um intervalo com várias células resultará em um erro sendo lançado (ou um intervalo `getSpillParent` nulo sendo retornado para `Range.getSpillParentOrNullObject` ).</span><span class="sxs-lookup"><span data-stu-id="010fb-117">Calling `getSpillParent` on a range with multiple cells will result in an error being thrown (or a null range being returned for `Range.getSpillParentOrNullObject`).</span></span>

## <a name="see-also"></a><span data-ttu-id="010fb-118">Confira também</span><span class="sxs-lookup"><span data-stu-id="010fb-118">See also</span></span>

- [<span data-ttu-id="010fb-119">Modelo de objeto JavaScript do Excel em Suplementos do Office</span><span class="sxs-lookup"><span data-stu-id="010fb-119">Excel JavaScript object model in Office Add-ins</span></span>](excel-add-ins-core-concepts.md)
- [<span data-ttu-id="010fb-120">Trabalhar com células usando a API JavaScript do Excel</span><span class="sxs-lookup"><span data-stu-id="010fb-120">Work with cells using the Excel JavaScript API</span></span>](excel-add-ins-cells.md)
- [<span data-ttu-id="010fb-121">Trabalhar simultaneamente com vários intervalos em suplementos do Excel</span><span class="sxs-lookup"><span data-stu-id="010fb-121">Work with multiple ranges simultaneously in Excel add-ins</span></span>](excel-add-ins-multiple-ranges.md)
