---
title: Valores em branco e nulos em suplementos do Excel
description: Saiba como trabalhar com um valor nulo em branco nos métodos e propriedades do modelo de objeto do Excel.
ms.date: 09/03/2020
localization_priority: Normal
ms.openlocfilehash: 3f38569f7342bb88c52ce424db426bfa7939be5e
ms.sourcegitcommit: c6308cf245ac1bc66a876eaa0a7bb4a2492991ac
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 09/08/2020
ms.locfileid: "47409376"
---
# <a name="blank-and-null-values-in-excel-add-ins"></a><span data-ttu-id="e96bd-103">Valores em branco e nulos em suplementos do Excel</span><span class="sxs-lookup"><span data-stu-id="e96bd-103">Blank and null values in Excel add-ins</span></span>

<span data-ttu-id="e96bd-104">`null` e as cadeias de caracteres esvaziadas têm implicações especiais nas APIs JavaScript do Excel.</span><span class="sxs-lookup"><span data-stu-id="e96bd-104">`null` and empty strings have special implications in the Excel JavaScript APIs.</span></span> <span data-ttu-id="e96bd-105">Elas são usadas para representar células vazias, sem formatação ou valores padrão.</span><span class="sxs-lookup"><span data-stu-id="e96bd-105">They're used to represent empty cells, no formatting, or default values.</span></span> <span data-ttu-id="e96bd-106">Essa seção detalha o uso da `null` e de uma cadeia de caracteres vazia ao obter e definir as propriedades.</span><span class="sxs-lookup"><span data-stu-id="e96bd-106">This section details the use of `null` and empty string when getting and setting properties.</span></span>

## <a name="null-input-in-2-d-array"></a><span data-ttu-id="e96bd-107">entrada nula em uma matriz 2D</span><span class="sxs-lookup"><span data-stu-id="e96bd-107">null input in 2-D Array</span></span>

<span data-ttu-id="e96bd-p102">No Excel, um intervalo é representado por uma matriz 2D, onde a primeira dimensão é linhas e a segunda dimensão é colunas. Para definir valores, o formato do número ou a fórmula apenas para células específicas em um intervalo, especifique os valores, o formato do número ou a fórmula para essas células na matriz 2D, bem como `null` para todas as outras células na matriz 2D.</span><span class="sxs-lookup"><span data-stu-id="e96bd-p102">In Excel, a range is represented by a 2-D array, where the first dimension is rows and the second dimension is columns. To set values, number format, or formula for only specific cells within a range, specify the values, number format, or formula for those cells in the 2-D array, and specify `null` for all other cells in the 2-D array.</span></span>

<span data-ttu-id="e96bd-p103">Por exemplo, para atualizar o formato do número apenas para uma célula em um intervalo e manter o formato de número existente para todas as outras células no intervalo, especifique o novo formato de número para a célula a ser atualizada e `null` para todas as outras células. O trecho de código a seguir define um novo formato de número para a quarta célula no intervalo e não altera o formato de número para as primeiras três células no intervalo.</span><span class="sxs-lookup"><span data-stu-id="e96bd-p103">For example, to update the number format for only one cell within a range, and retain the existing number format for all other cells in the range, specify the new number format for the cell to update, and specify `null` for all other cells. The following code snippet sets a new number format for the fourth cell in the range, and leaves the number format unchanged for the first three cells in the range.</span></span>

```js
range.values = [['Eurasia', '29.96', '0.25', '15-Feb' ]];
range.numberFormat = [[null, null, null, 'm/d/yyyy;@']];
```

## <a name="null-input-for-a-property"></a><span data-ttu-id="e96bd-112">entrada nula para uma propriedade</span><span class="sxs-lookup"><span data-stu-id="e96bd-112">null input for a property</span></span>

<span data-ttu-id="e96bd-p104">`null` não é uma entrada válida para uma propriedade única. Por exemplo, o trecho de código a seguir não é válido, pois a propriedade `values` do intervalo não pode ser definida como `null`.</span><span class="sxs-lookup"><span data-stu-id="e96bd-p104">`null` is not a valid input for single property. For example, the following code snippet is not valid, as the `values` property of the range cannot be set to `null`.</span></span>

```js
range.values = null; // This is not a valid snippet. 
```

<span data-ttu-id="e96bd-115">Da mesma forma, o seguinte snippet de código não é válido, pois `null` não é um valor válido para a propriedade `color`.</span><span class="sxs-lookup"><span data-stu-id="e96bd-115">Likewise, the following code snippet is not valid, as `null` is not a valid value for the `color` property.</span></span>

```js
range.format.fill.color =  null;  // This is not a valid snippet. 
```

## <a name="null-property-values-in-the-response"></a><span data-ttu-id="e96bd-116">Valores da propriedade nula na resposta</span><span class="sxs-lookup"><span data-stu-id="e96bd-116">null property values in the response</span></span>

<span data-ttu-id="e96bd-p105">A formatação de propriedades como `size` e `color` conterá valores `null` na resposta quando valores diferentes existirem no intervalo especificado. Por exemplo, se você recuperar um intervalo e carregar sua propriedade `format.font.color`:</span><span class="sxs-lookup"><span data-stu-id="e96bd-p105">Formatting properties such as `size` and `color` will contain `null` values in the response when different values exist in the specified range. For example, if you retrieve a range and load its `format.font.color` property:</span></span>

* <span data-ttu-id="e96bd-119">Se todas as células no intervalo tiverem a mesma cor de fonte, `range.format.font.color` especificará essa cor.</span><span class="sxs-lookup"><span data-stu-id="e96bd-119">If all cells in the range have the same font color, `range.format.font.color` specifies that color.</span></span>
* <span data-ttu-id="e96bd-120">Se houver várias cores de fonte dentro do intervalo, `range.format.font.color` será `null`.</span><span class="sxs-lookup"><span data-stu-id="e96bd-120">If multiple font colors are present within the range, `range.format.font.color` is `null`.</span></span>

## <a name="blank-input-for-a-property"></a><span data-ttu-id="e96bd-121">Entrada em branco para uma propriedade</span><span class="sxs-lookup"><span data-stu-id="e96bd-121">Blank input for a property</span></span>

<span data-ttu-id="e96bd-p106">Quando você especificar um valor em branco para uma propriedade (isto é, duas aspas sem espaço entre elas `''`), ele será interpretado como uma instrução para limpar ou redefinir a propriedade. Por exemplo:</span><span class="sxs-lookup"><span data-stu-id="e96bd-p106">When you specify a blank value for a property (i.e., two quotation marks with no space in-between `''`), it will be interpreted as an instruction to clear or reset the property. For example:</span></span>

* <span data-ttu-id="e96bd-124">Se você especificar um valor em branco para a propriedade `values` de um intervalo, o conteúdo do intervalo será apagado.</span><span class="sxs-lookup"><span data-stu-id="e96bd-124">If you specify a blank value for the `values` property of a range, the content of the range is cleared.</span></span>
* <span data-ttu-id="e96bd-125">Se você especificar um valor em branco para a propriedade `numberFormat`, o formato de número será redefinido para `General`.</span><span class="sxs-lookup"><span data-stu-id="e96bd-125">If you specify a blank value for the `numberFormat` property, the number format is reset to `General`.</span></span>
* <span data-ttu-id="e96bd-126">Se você especificar um valor em branco para a propriedade `formula` e a propriedade `formulaLocale`, os valores de fórmula serão apagados.</span><span class="sxs-lookup"><span data-stu-id="e96bd-126">If you specify a blank value for the `formula` property and `formulaLocale` property, the formula values are cleared.</span></span>

## <a name="blank-property-values-in-the-response"></a><span data-ttu-id="e96bd-127">Valores da propriedade em branco na resposta</span><span class="sxs-lookup"><span data-stu-id="e96bd-127">Blank property values in the response</span></span>

<span data-ttu-id="e96bd-p107">Para operações de leitura, um valor de propriedade em branco na resposta (isto é, duas aspas sem espaço entre elas `''`) indica que a célula não contém dados nem valor. No primeiro exemplo abaixo, a primeira e a última célula no intervalo não contêm dados. No segundo exemplo, as primeiras duas células no intervalo não contêm uma fórmula.</span><span class="sxs-lookup"><span data-stu-id="e96bd-p107">For read operations, a blank property value in the response (i.e., two quotation marks with no space in-between `''`) indicates that cell contains no data or value. In the first example below, the first and last cell in the range contain no data. In the second example, the first two cells in the range do not contain a formula.</span></span>

```js
range.values = [['', 'some', 'data', 'in', 'other', 'cells', '']];
```

```js
range.formula = [['', '', '=Rand()']];
```
