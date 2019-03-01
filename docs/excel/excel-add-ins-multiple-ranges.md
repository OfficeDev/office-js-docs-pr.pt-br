---
title: Trabalhar simultaneamente com vários intervalos em suplementos do Excel
description: ''
ms.date: 02/20/2019
localization_priority: Normal
ms.openlocfilehash: c6bbbaee6f6cbfda5d495f533caf3dbe1325401b
ms.sourcegitcommit: 8e20e7663be2aaa0f7a5436a965324d171bc667d
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 02/28/2019
ms.locfileid: "30199603"
---
# <a name="work-with-multiple-ranges-simultaneously-in-excel-add-ins-preview"></a><span data-ttu-id="4a76d-102">Trabalhar simultaneamente com vários intervalos em suplementos do Excel (Visualização)</span><span class="sxs-lookup"><span data-stu-id="4a76d-102">Work with multiple ranges simultaneously in Excel add-ins (preview)</span></span>

<span data-ttu-id="4a76d-103">A biblioteca de JavaScript do Excel permite que o suplemento realize operações e defina propriedades, em vários intervalos simultaneamente.</span><span class="sxs-lookup"><span data-stu-id="4a76d-103">The Excel JavaScript library enables your add-in to perform operations, and set properties, on multiple ranges simultaneously.</span></span> <span data-ttu-id="4a76d-104">Os intervalos não precisam ser contíguos.</span><span class="sxs-lookup"><span data-stu-id="4a76d-104">The ranges do not have to be contiguous.</span></span> <span data-ttu-id="4a76d-105">Além de tornar seu código mais simples, essa maneira de definir uma propriedade é executada muito mais rapidamente do que definir a mesma propriedade individualmente para cada um dos intervalos.</span><span class="sxs-lookup"><span data-stu-id="4a76d-105">In addition to making your code simpler, this way of setting a property runs much faster than setting the same property individually for each of the ranges.</span></span>

> [!NOTE]
> <span data-ttu-id="4a76d-106">As APIs descritas neste artigo requerem a \*\* versão 1809 Build 10820.20000 clique para executar do Office 2016\*\* ou posterior.</span><span class="sxs-lookup"><span data-stu-id="4a76d-106">The APIs described in this article require **Office 2016 Click-to-Run version 1809 Build 10820.20000** or later.</span></span> <span data-ttu-id="4a76d-107">(Talvez seja necessário participar do [programa Office](https://products.office.com/office-insider) Insider para obter uma compilação apropriada.)[!INCLUDE [Information about using preview APIs](../includes/using-preview-apis.md)]</span><span class="sxs-lookup"><span data-stu-id="4a76d-107">(You may need to join the [Office Insider program](https://products.office.com/office-insider) to get an appropriate build.)  [!INCLUDE [Information about using preview APIs](../includes/using-preview-apis.md)]</span></span>

## <a name="rangeareas"></a><span data-ttu-id="4a76d-108">RangeAreas</span><span class="sxs-lookup"><span data-stu-id="4a76d-108">RangeAreas</span></span>

<span data-ttu-id="4a76d-109">Um conjunto de intervalos (possivelmente não contíguos) é representado por um objeto [RangeAreas](/javascript/api/excel/excel.rangeareas) .</span><span class="sxs-lookup"><span data-stu-id="4a76d-109">A set of (possibly discontiguous) ranges is represented by a [RangeAreas](/javascript/api/excel/excel.rangeareas) object.</span></span> <span data-ttu-id="4a76d-110">Possui propriedades e métodos semelhantes ao tipo `Range` (muitos com os mesmos nomes ou semelhantes), mas foram feitos ajustes para:</span><span class="sxs-lookup"><span data-stu-id="4a76d-110">It has properties and methods similar to the `Range` type (many with the same, or similar, names), but adjustments have been made to:</span></span>

- <span data-ttu-id="4a76d-111">Os tipos de dados para propriedades e o comportamento dos setters e getters.</span><span class="sxs-lookup"><span data-stu-id="4a76d-111">The data types for properties and the behavior of the setters and getters.</span></span>
- <span data-ttu-id="4a76d-112">Os tipos de dados dos parâmetros do método e os comportamentos do método.</span><span class="sxs-lookup"><span data-stu-id="4a76d-112">The data types of method parameters and the method behaviors.</span></span>
- <span data-ttu-id="4a76d-113">Os tipos de dados de forma retornam valores.</span><span class="sxs-lookup"><span data-stu-id="4a76d-113">The data types of method return values.</span></span>

<span data-ttu-id="4a76d-114">Alguns exemplos:</span><span class="sxs-lookup"><span data-stu-id="4a76d-114">Some examples:</span></span>

- <span data-ttu-id="4a76d-115">`RangeAreas` tem uma propriedade `address` que retorna uma cadeia de caracteres delimitada por vírgula de intervalo de endereços, em vez de apenas um endereço como na propriedade`Range.address`.</span><span class="sxs-lookup"><span data-stu-id="4a76d-115">`RangeAreas` has an `address` property that returns a comma-delimited string of range addresses, instead of just one address as with the `Range.address` property.</span></span>
- <span data-ttu-id="4a76d-116">`RangeAreas` tem uma propriedade `dataValidation` que retorna um objeto `DataValidation` que representa a validação de dados de todos os intervalos em`RangeAreas`, se for consistente.</span><span class="sxs-lookup"><span data-stu-id="4a76d-116">`RangeAreas` has a `dataValidation` property that returns a `DataValidation` object that represents the data validation of all the ranges in the `RangeAreas`, if it is consistent.</span></span> <span data-ttu-id="4a76d-117">A propriedade é `null` se objetos idênticos `DataValidation` não forem aplicados a todos os intervalos em `RangeAreas`.</span><span class="sxs-lookup"><span data-stu-id="4a76d-117">The property is `null` if identical `DataValidation` objects are not applied to all the all the ranges in the `RangeAreas`.</span></span> <span data-ttu-id="4a76d-118">Esse é um princípio geral, mas não universal com o objeto `RangeAreas`: *se uma propriedade não têm valores consistentes em todos os todos os intervalos em `RangeAreas`, então será `null`.*</span><span class="sxs-lookup"><span data-stu-id="4a76d-118">This is a general, but not universal, principle with the `RangeAreas` object: *If a property does not have consistent values on all the all the ranges in the `RangeAreas`, then it is `null`.*</span></span> <span data-ttu-id="4a76d-119">Ver [ler as propriedades de RangeAreas](#read-properties-of-rangeareas) para mais informações e algumas exceções.</span><span class="sxs-lookup"><span data-stu-id="4a76d-119">See [Read properties of RangeAreas](#read-properties-of-rangeareas) for more information and some exceptions.</span></span>
- <span data-ttu-id="4a76d-120">`RangeAreas.cellCount` é o número total de células em todos os intervalos no `RangeAreas`.</span><span class="sxs-lookup"><span data-stu-id="4a76d-120">`RangeAreas.cellCount` gets the total number of cells in all the ranges in the `RangeAreas`.</span></span>
- <span data-ttu-id="4a76d-121">`RangeAreas.calculate` recalcula as células de todos os intervalos no `RangeAreas`.</span><span class="sxs-lookup"><span data-stu-id="4a76d-121">`RangeAreas.calculate` recalculates the cells of all the ranges in the `RangeAreas`.</span></span>
- <span data-ttu-id="4a76d-122">`RangeAreas.getEntireColumn` e `RangeAreas.getEntireRow` retornar outra `RangeAreas` objeto que representa todas as colunas (ou linhas) em todos os intervalos no `RangeAreas`.</span><span class="sxs-lookup"><span data-stu-id="4a76d-122">`RangeAreas.getEntireColumn` and `RangeAreas.getEntireRow` return another `RangeAreas` object that represents all of the columns (or rows) in all the ranges in the `RangeAreas`.</span></span> <span data-ttu-id="4a76d-123">Por exemplo, se `RangeAreas` representa "A1: C4" e "F14:L15" em seguida, `RangeAreas.getEntireColumn` retorna um objeto `RangeAreas` que representa "A:C" e "F:L".</span><span class="sxs-lookup"><span data-stu-id="4a76d-123">For example, if the `RangeAreas` represents "A1:C4" and "F14:L15", then `RangeAreas.getEntireColumn` returns a `RangeAreas` object that represents "A:C" and "F:L".</span></span>
- <span data-ttu-id="4a76d-124">`RangeAreas.copyFrom` pode ter o parâmetro `Range` ou `RangeAreas` que representam os intervalos de origem da operação de cópia.</span><span class="sxs-lookup"><span data-stu-id="4a76d-124">`RangeAreas.copyFrom` can take either a `Range` or a `RangeAreas` parameter representing the source range(s) of the copy operation.</span></span>

#### <a name="complete-list-of-range-members-that-are-also-available-on-rangeareas"></a><span data-ttu-id="4a76d-125">Lista completa de membros do intervalo que também estão disponíveis em RangeAreas</span><span class="sxs-lookup"><span data-stu-id="4a76d-125">Complete list of Range members that are also available on RangeAreas</span></span>

##### <a name="properties"></a><span data-ttu-id="4a76d-126">Propriedades</span><span class="sxs-lookup"><span data-stu-id="4a76d-126">Properties</span></span>

<span data-ttu-id="4a76d-127">Familiarize-se com as [Propriedades de leitura do RangeAreas](#read-properties-of-rangeareas) antes de escrever o código que lê as propriedades listadas.</span><span class="sxs-lookup"><span data-stu-id="4a76d-127">Be familiar with [Read properties of RangeAreas](#read-properties-of-rangeareas) before you write code that reads any properties listed.</span></span> <span data-ttu-id="4a76d-128">Existem sutilezas para o que é retornado.</span><span class="sxs-lookup"><span data-stu-id="4a76d-128">There are subtleties to what gets returned.</span></span>

- <span data-ttu-id="4a76d-129">address</span><span class="sxs-lookup"><span data-stu-id="4a76d-129">address</span></span>
- <span data-ttu-id="4a76d-130">addressLocal</span><span class="sxs-lookup"><span data-stu-id="4a76d-130">addressLocal</span></span>
- <span data-ttu-id="4a76d-131">cellCount</span><span class="sxs-lookup"><span data-stu-id="4a76d-131">cellCount</span></span>
- <span data-ttu-id="4a76d-132">conditionalFormats</span><span class="sxs-lookup"><span data-stu-id="4a76d-132">conditionalFormats</span></span>
- <span data-ttu-id="4a76d-133">context</span><span class="sxs-lookup"><span data-stu-id="4a76d-133">context</span></span>
- <span data-ttu-id="4a76d-134">dataValidation</span><span class="sxs-lookup"><span data-stu-id="4a76d-134">dataValidation</span></span>
- <span data-ttu-id="4a76d-135">formato</span><span class="sxs-lookup"><span data-stu-id="4a76d-135">format</span></span>
- <span data-ttu-id="4a76d-136">isEntireColumn</span><span class="sxs-lookup"><span data-stu-id="4a76d-136">isEntireColumn</span></span>
- <span data-ttu-id="4a76d-137">isEntireRow</span><span class="sxs-lookup"><span data-stu-id="4a76d-137">isEntireRow</span></span>
- <span data-ttu-id="4a76d-138">style</span><span class="sxs-lookup"><span data-stu-id="4a76d-138">style</span></span>
- <span data-ttu-id="4a76d-139">planilha</span><span class="sxs-lookup"><span data-stu-id="4a76d-139">worksheet</span></span>

##### <a name="methods"></a><span data-ttu-id="4a76d-140">Métodos</span><span class="sxs-lookup"><span data-stu-id="4a76d-140">Methods</span></span>

<span data-ttu-id="4a76d-141">Os métodos de intervalo na visualização são marcados.</span><span class="sxs-lookup"><span data-stu-id="4a76d-141">Range methods in preview are marked.</span></span>

- <span data-ttu-id="4a76d-142">calculate()</span><span class="sxs-lookup"><span data-stu-id="4a76d-142">calculate()</span></span>
- <span data-ttu-id="4a76d-143">clear()</span><span class="sxs-lookup"><span data-stu-id="4a76d-143">clear()</span></span>
- <span data-ttu-id="4a76d-144">convertDataTypeToText() (visualização)</span><span class="sxs-lookup"><span data-stu-id="4a76d-144">convertDataTypeToText() (preview)</span></span>
- <span data-ttu-id="4a76d-145">convertToLinkedDataType() (visualização)</span><span class="sxs-lookup"><span data-stu-id="4a76d-145">convertToLinkedDataType() (preview)</span></span>
- <span data-ttu-id="4a76d-146">copyFrom() (visualização)</span><span class="sxs-lookup"><span data-stu-id="4a76d-146">copyFrom() (preview)</span></span>
- <span data-ttu-id="4a76d-147">getEntireColumn()</span><span class="sxs-lookup"><span data-stu-id="4a76d-147">getEntireColumn()</span></span>
- <span data-ttu-id="4a76d-148">getEntireRow()</span><span class="sxs-lookup"><span data-stu-id="4a76d-148">getEntireRow()</span></span>
- <span data-ttu-id="4a76d-149">getIntersection()</span><span class="sxs-lookup"><span data-stu-id="4a76d-149">getIntersection()</span></span>
- <span data-ttu-id="4a76d-150">getIntersectionOrNullObject()</span><span class="sxs-lookup"><span data-stu-id="4a76d-150">getIntersectionOrNullObject()</span></span>
- <span data-ttu-id="4a76d-151">getOffsetRange() (chamada getOffsetRangeAreas no objeto RangeAreas)</span><span class="sxs-lookup"><span data-stu-id="4a76d-151">getOffsetRange() (named getOffsetRangeAreas on the RangeAreas object)</span></span>
- <span data-ttu-id="4a76d-152">getSpecialCells() (visualização)</span><span class="sxs-lookup"><span data-stu-id="4a76d-152">getSpecialCells() (preview)</span></span>
- <span data-ttu-id="4a76d-153">getSpecialCellsOrNullObject() (visualização)</span><span class="sxs-lookup"><span data-stu-id="4a76d-153">getSpecialCellsOrNullObject() (preview)</span></span>
- <span data-ttu-id="4a76d-154">getTables() (visualização)</span><span class="sxs-lookup"><span data-stu-id="4a76d-154">getTables() (preview)</span></span>
- <span data-ttu-id="4a76d-155">getUsedRange() (chamada getUsedRangeAreas no objeto RangeAreas)</span><span class="sxs-lookup"><span data-stu-id="4a76d-155">getUsedRange() (named getUsedRangeAreas on the RangeAreas object)</span></span>
- <span data-ttu-id="4a76d-156">getUsedRangeOrNullObject() (chamada getUsedRangeAreasOrNullObject no objeto RangeAreas)</span><span class="sxs-lookup"><span data-stu-id="4a76d-156">getUsedRangeOrNullObject() (named getUsedRangeAreasOrNullObject on the RangeAreas object)</span></span>
- <span data-ttu-id="4a76d-157">load()</span><span class="sxs-lookup"><span data-stu-id="4a76d-157">load()</span></span>
- <span data-ttu-id="4a76d-158">set()</span><span class="sxs-lookup"><span data-stu-id="4a76d-158">set()</span></span>
- <span data-ttu-id="4a76d-159">setDirty() (visualização)</span><span class="sxs-lookup"><span data-stu-id="4a76d-159">setDirty() (preview)</span></span>
- <span data-ttu-id="4a76d-160">toJSON()</span><span class="sxs-lookup"><span data-stu-id="4a76d-160">toJSON()</span></span>
- <span data-ttu-id="4a76d-161">track()</span><span class="sxs-lookup"><span data-stu-id="4a76d-161">track()</span></span>
- <span data-ttu-id="4a76d-162">untrack()</span><span class="sxs-lookup"><span data-stu-id="4a76d-162">untrack()</span></span>

### <a name="rangearea-specific-properties-and-methods"></a><span data-ttu-id="4a76d-163">Métodos e propriedades específicos do RangeArea</span><span class="sxs-lookup"><span data-stu-id="4a76d-163">RangeArea-specific properties and methods</span></span>

<span data-ttu-id="4a76d-164">O tipo `RangeAreas` tem alguns métodos e propriedades que não estão no objeto `Range`.</span><span class="sxs-lookup"><span data-stu-id="4a76d-164">The `RangeAreas` type has some properties and methods that are not on the `Range` object.</span></span> <span data-ttu-id="4a76d-165">Esta é a seleção deles:</span><span class="sxs-lookup"><span data-stu-id="4a76d-165">The following is a selection of them:</span></span>

- <span data-ttu-id="4a76d-166">`areas`: O objeto `RangeCollection` que contém todos os intervalos representados pelo objeto `RangeAreas`.</span><span class="sxs-lookup"><span data-stu-id="4a76d-166">`areas`: A `RangeCollection` object that contains all of the ranges represented by the `RangeAreas` object.</span></span> <span data-ttu-id="4a76d-167">O objeto `RangeCollection` também é novidade e é semelhante a outros objetos do conjunto do Excel.</span><span class="sxs-lookup"><span data-stu-id="4a76d-167">The `RangeCollection` object is also new and is similar to other Excel collection objects.</span></span> <span data-ttu-id="4a76d-168">É uma propriedade `items` que é uma matriz de objetos `Range` que representam os intervalos.</span><span class="sxs-lookup"><span data-stu-id="4a76d-168">It has an `items` property which is an array of `Range` objects representing the ranges.</span></span>
- <span data-ttu-id="4a76d-169">`areaCount`: O número total de intervalos em `RangeAreas`.</span><span class="sxs-lookup"><span data-stu-id="4a76d-169">`areaCount`: The total number of ranges in the `RangeAreas`.</span></span>
- <span data-ttu-id="4a76d-170">`getOffsetRangeAreas`: Funciona como [Range.getOffsetRange](/javascript/api/excel/excel.range#getoffsetrange-rowoffset--columnoffset-), exceto pelo fato de que o `RangeAreas` é retornado e contém os intervalos que são todos os deslocamentos de um dos intervalos do `RangeAreas` original.</span><span class="sxs-lookup"><span data-stu-id="4a76d-170">`getOffsetRangeAreas`: Works just like [Range.getOffsetRange](/javascript/api/excel/excel.range#getoffsetrange-rowoffset--columnoffset-), except that a `RangeAreas` is returned and it contains ranges that are each offset from one of the ranges in the original `RangeAreas`.</span></span>

## <a name="create-rangeareas"></a><span data-ttu-id="4a76d-171">Criar RangeAreas</span><span class="sxs-lookup"><span data-stu-id="4a76d-171">Create RangeAreas</span></span>

<span data-ttu-id="4a76d-172">Você pode criar o objeto`RangeAreas` de duas maneiras básicas:</span><span class="sxs-lookup"><span data-stu-id="4a76d-172">You can create `RangeAreas` object in two basic ways:</span></span>

- <span data-ttu-id="4a76d-173">Ligue `Worksheet.getRanges()` e encaminhe-o em uma cadeia de caracteres com endereços de intervalo separado por vírgula.</span><span class="sxs-lookup"><span data-stu-id="4a76d-173">Call `Worksheet.getRanges()` and pass it a string with comma-delimited range addresses.</span></span> <span data-ttu-id="4a76d-174">Se algum intervalo que você deseja incluir tiver sido feito em um [NamedItem](/javascript/api/excel/excel.nameditem), você poderá incluir o nome, em vez do endereço, cadeia de caracteres.</span><span class="sxs-lookup"><span data-stu-id="4a76d-174">If any range you want to include has been made into a [NamedItem](/javascript/api/excel/excel.nameditem), you can include the name, instead of the address, in the string.</span></span>
- <span data-ttu-id="4a76d-175">Chamar `Workbook.getSelectedRanges()`.</span><span class="sxs-lookup"><span data-stu-id="4a76d-175">Call `Workbook.getSelectedRanges()`.</span></span> <span data-ttu-id="4a76d-176">Esse método retornará um `RangeAreas` representando todos os intervalos selecionados na planilha ativa no momento.</span><span class="sxs-lookup"><span data-stu-id="4a76d-176">This method returns a `RangeAreas` representing all the ranges that are selected on the currently active worksheet.</span></span>

<span data-ttu-id="4a76d-177">Quando você tiver um objeto `RangeAreas`, você pode criar outros usando os métodos de objeto que retornam `RangeAreas` como `getOffsetRangeAreas` e `getIntersection`.</span><span class="sxs-lookup"><span data-stu-id="4a76d-177">Once you have a `RangeAreas` object, you can create others using the methods on the object that return `RangeAreas` such as `getOffsetRangeAreas` and `getIntersection`.</span></span>

> [!NOTE]
> <span data-ttu-id="4a76d-178">É possível adicionar diretamente intervalos adicionais para um objeto `RangeAreas`.</span><span class="sxs-lookup"><span data-stu-id="4a76d-178">You cannot directly add additional ranges to a `RangeAreas` object.</span></span> <span data-ttu-id="4a76d-179">Por exemplo, o conjunto `RangeAreas.areas` não tem um método`add`.</span><span class="sxs-lookup"><span data-stu-id="4a76d-179">For example, the collection in `RangeAreas.areas` does not have an `add` method.</span></span>

> [!WARNING]
> <span data-ttu-id="4a76d-180">Tente adicionar ou excluir membros diretamente à matriz`RangeAreas.areas.items`.</span><span class="sxs-lookup"><span data-stu-id="4a76d-180">Do not attempt to directly add or delete members of the the `RangeAreas.areas.items` array.</span></span> <span data-ttu-id="4a76d-181">Isso levará a um comportamento indesejável no seu código.</span><span class="sxs-lookup"><span data-stu-id="4a76d-181">This will lead to undesirable behavior in your code.</span></span> <span data-ttu-id="4a76d-182">Por exemplo, é possível enviar um objeto adicional `Range` para a matriz, mas isso causará erros porque as propriedades e métodos `RangeAreas` se comportam como se o novo item não estivesse ali.</span><span class="sxs-lookup"><span data-stu-id="4a76d-182">For example, it is possible to push an additional `Range` object onto the array, but doing so will cause errors because `RangeAreas` properties and methods behave as if the new item isn't there.</span></span> <span data-ttu-id="4a76d-183">Por exemplo, a propriedade `areaCount` não inclui intervalos transferidos dessa maneira e o `RangeAreas.getItemAt(index)` gera um erro se `index` for maior que `areasCount-1`.</span><span class="sxs-lookup"><span data-stu-id="4a76d-183">For example, the `areaCount` property does not include ranges pushed in this way, and the `RangeAreas.getItemAt(index)` throws an error if `index` is larger than `areasCount-1`.</span></span> <span data-ttu-id="4a76d-184">Da mesma forma, excluir um objeto `Range` na matriz `RangeAreas.areas.items` obtendo uma referência a ele e chamando seu método `Range.delete` causa bugs: embora o `Range`objeto\* seja \*excluído, as propriedades e métodos do objeto pai `RangeAreas` se comportam ou tentam se comportar, como se ele ainda existisse.</span><span class="sxs-lookup"><span data-stu-id="4a76d-184">Similarly, deleting a `Range` object in the `RangeAreas.areas.items` array by getting a reference to it and calling its `Range.delete` method causes bugs: although the `Range` object *is* deleted, the properties and methods of the parent `RangeAreas` object behave, or try to, as if it is still in existence.</span></span> <span data-ttu-id="4a76d-185">Por exemplo, se o seu código chamar `RangeAreas.calculate`, o Office tentará calcular o intervalo, mas haverá erro porque o objeto de intervalo desapareceu.</span><span class="sxs-lookup"><span data-stu-id="4a76d-185">For example, if your code calls `RangeAreas.calculate`, Office will try to calculate the range, but will error because the range object is gone.</span></span>

## <a name="set-properties-on-multiple-ranges"></a><span data-ttu-id="4a76d-186">Definir as propriedades em vários intervalos</span><span class="sxs-lookup"><span data-stu-id="4a76d-186">Set properties on multiple ranges</span></span>

<span data-ttu-id="4a76d-187">Definir uma propriedade em um `RangeAreas` objeto define a propriedade correspondente em todos os intervalos no conjunto `RangeAreas.areas`.</span><span class="sxs-lookup"><span data-stu-id="4a76d-187">Setting a property on a `RangeAreas` object sets the corresponding property on all the ranges in the `RangeAreas.areas` collection.</span></span>

<span data-ttu-id="4a76d-188">A seguir, um exemplo de configuração de uma propriedade em vários intervalos.</span><span class="sxs-lookup"><span data-stu-id="4a76d-188">The following is an example of setting a property on multiple ranges.</span></span> <span data-ttu-id="4a76d-189">A função realça os intervalos **F3:F5** e **H3:H5**.</span><span class="sxs-lookup"><span data-stu-id="4a76d-189">The function highlights the ranges **F3:F5** and **H3:H5**.</span></span>

```js
Excel.run(function (context) {
    var sheet = context.workbook.worksheets.getActiveWorksheet();
    var rangeAreas = sheet.getRanges("F3:F5, H3:H5");
    rangeAreas.format.fill.color = "pink";

    return context.sync();
})
```

<span data-ttu-id="4a76d-190">Este exemplo se aplica a cenários nos quais você pode codificar os endereços de intervalo para os quais você passa para `getRanges` ou facilmente calculá-los no tempo de execução.</span><span class="sxs-lookup"><span data-stu-id="4a76d-190">This example applies to scenarios in which you can hard code the range addresses that you pass to `getRanges` or easily calculate them at runtime.</span></span> <span data-ttu-id="4a76d-191">Alguns dos cenários em que isso pode ser verdadeiro incluem:</span><span class="sxs-lookup"><span data-stu-id="4a76d-191">Some of the scenarios in which this would be true include:</span></span>

- <span data-ttu-id="4a76d-192">O código é executado no contexto de um modelo conhecido.</span><span class="sxs-lookup"><span data-stu-id="4a76d-192">The code runs in the context of a known template.</span></span>
- <span data-ttu-id="4a76d-193">O código é executado no contexto de dados importados, em que o esquema dos dados é conhecido.</span><span class="sxs-lookup"><span data-stu-id="4a76d-193">The code runs in the context of imported data where the schema of the data is known.</span></span>

## <a name="get-special-cells-from-multiple-ranges"></a><span data-ttu-id="4a76d-194">Obter células especiais de vários intervalos</span><span class="sxs-lookup"><span data-stu-id="4a76d-194">Get special cells from multiple ranges</span></span>

<span data-ttu-id="4a76d-195">As `getSpecialCells` e `getSpecialCellsOrNullObject` métodos no `RangeAreas` objeto funciona analogamente para métodos de mesmo nome no `Range` objeto.</span><span class="sxs-lookup"><span data-stu-id="4a76d-195">The `getSpecialCells` and `getSpecialCellsOrNullObject` methods on the `RangeAreas` object work analogously to methods of the same name on the `Range` object.</span></span> <span data-ttu-id="4a76d-196">Esses métodos retornam as células com característica especificada de todos os intervalos no `RangeAreas.areas` conjunto.</span><span class="sxs-lookup"><span data-stu-id="4a76d-196">These methods return the cells with the specified characteristic from all of the ranges in the `RangeAreas.areas` collection.</span></span> <span data-ttu-id="4a76d-197">Confira a seção [Localizar células especiais em um intervalo](excel-add-ins-ranges-advanced.md#find-special-cells-within-a-range-preview) para saber mais sobre células especiais.</span><span class="sxs-lookup"><span data-stu-id="4a76d-197">See the [Find special cells within a range](excel-add-ins-ranges-advanced.md#find-special-cells-within-a-range-preview) section for more details on special cells.</span></span>

<span data-ttu-id="4a76d-198">Ao chamar as `getSpecialCells` ou `getSpecialCellsOrNullObject` método em um `RangeAreas` objeto:</span><span class="sxs-lookup"><span data-stu-id="4a76d-198">When calling the `getSpecialCells` or `getSpecialCellsOrNullObject` method on a `RangeAreas` object:</span></span>

- <span data-ttu-id="4a76d-199">Se você passar `Excel.SpecialCellType.sameConditionalFormat` como o primeiro parâmetro, o método retorna todas as células com a mesma formatação condicional que a célula superior esquerda do primeiro intervalo no `RangeAreas.areas` conjunto.</span><span class="sxs-lookup"><span data-stu-id="4a76d-199">If you pass `Excel.SpecialCellType.sameConditionalFormat` as the first parameter, the method returns all cells with the same conditional formatting as the upper-leftmost cell in the first range in the `RangeAreas.areas` collection.</span></span>
- <span data-ttu-id="4a76d-200">Se você passar `Excel.SpecialCellType.sameDataValidation` como o primeiro parâmetro, o método retorna todas as células com a regra de validação de dados que a célula superior esquerda do primeiro intervalo no `RangeAreas.areas` conjunto.</span><span class="sxs-lookup"><span data-stu-id="4a76d-200">If you pass `Excel.SpecialCellType.sameDataValidation` as the first parameter, the method returns all cells with the same data validation rule as the upper-leftmost cell in the first range in the `RangeAreas.areas` collection.</span></span>

## <a name="read-properties-of-rangeareas"></a><span data-ttu-id="4a76d-201">Ler propriedades de RangeAreas</span><span class="sxs-lookup"><span data-stu-id="4a76d-201">Read properties of RangeAreas</span></span>

<span data-ttu-id="4a76d-202">A leitura de valores de propriedade `RangeAreas` requer cuidados, porque uma determinada propriedade pode ter valores diferentes para intervalos diferentes dentro de`RangeAreas`.</span><span class="sxs-lookup"><span data-stu-id="4a76d-202">Reading property values of `RangeAreas` requires care, because a given property may have different values for different ranges within the `RangeAreas`.</span></span> <span data-ttu-id="4a76d-203">A regra geral é que, se um valor consistente *puder* ser retornado, ele será retornado.</span><span class="sxs-lookup"><span data-stu-id="4a76d-203">The general rule is that if a consistent value *can* be returned it will be returned.</span></span> <span data-ttu-id="4a76d-204">Por exemplo, no código a seguir, O código RGB para pink (`#FFC0CB`) e `true` será registrado no console porque ambos os intervalos no objeto `RangeAreas` têm um preenchimento rosa e ambos são colunas inteiras.</span><span class="sxs-lookup"><span data-stu-id="4a76d-204">For example, in the following code, The RGB code for pink (`#FFC0CB`) and `true` will be logged to the console because both the ranges in the `RangeAreas` object have a pink fill and both are entire columns.</span></span>

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

<span data-ttu-id="4a76d-205">As coisas ficam mais complicadas quando a consistência não é possível.</span><span class="sxs-lookup"><span data-stu-id="4a76d-205">Things get more complicated when consistency isn't possible.</span></span> <span data-ttu-id="4a76d-206">O comportamento das propriedades `RangeAreas` seguem estes três princípios de três:</span><span class="sxs-lookup"><span data-stu-id="4a76d-206">The behavior of `RangeAreas` properties follows these three principles:</span></span>

- <span data-ttu-id="4a76d-207">Uma propriedade booliana de um `RangeAreas`retorno de objeto `false`, a menos que a propriedade seja verdadeira para todos os intervalos de membro.</span><span class="sxs-lookup"><span data-stu-id="4a76d-207">A boolean property of a `RangeAreas` object returns `false` unless the property is true for all the member ranges.</span></span>
- <span data-ttu-id="4a76d-208">Propriedades não boolianas, com exceção da propriedade `address`, retornam `null`, a menos que a propriedade correspondente em todos os intervalos de membros tenha o mesmo valor.</span><span class="sxs-lookup"><span data-stu-id="4a76d-208">Non-boolean properties, with the exception of the `address` property, return `null` unless the corresponding property on all the member ranges has the same value.</span></span>
- <span data-ttu-id="4a76d-209">A propriedade `address` retorna uma cadeia de caracteres delimitada por vírgulas dos endereços e intervalos dos membros.</span><span class="sxs-lookup"><span data-stu-id="4a76d-209">The `address` property returns a comma-delimited string of the addresses of the member ranges.</span></span>

<span data-ttu-id="4a76d-210">Por exemplo, o código a seguir cria um `RangeAreas` no qual apenas um intervalo é uma coluna inteira e apenas um é preenchido com rosa.</span><span class="sxs-lookup"><span data-stu-id="4a76d-210">For example, the following code creates a `RangeAreas` in which only one range is an entire column and only one is filled with pink.</span></span> <span data-ttu-id="4a76d-211">O console mostrará `null` para a cor de preenchimento `false` para a propriedade `isEntireRow` e "Planilha1! F3:F5, Planilha1! H:H"(supondo que o nome da planilha  seja "Planilha1") para a propriedade`address`.</span><span class="sxs-lookup"><span data-stu-id="4a76d-211">The console will show `null` for the fill color, `false` for the `isEntireRow` property, and "Sheet1!F3:F5, Sheet1!H:H" (assuming the sheet name is "Sheet1") for the `address` property.</span></span>

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

## <a name="see-also"></a><span data-ttu-id="4a76d-212">Confira também</span><span class="sxs-lookup"><span data-stu-id="4a76d-212">See also</span></span>

- [<span data-ttu-id="4a76d-213">Conceitos fundamentais de programação com a API JavaScript do Excel</span><span class="sxs-lookup"><span data-stu-id="4a76d-213">Fundamental programming concepts with the Excel JavaScript API</span></span>](../reference/overview/excel-add-ins-reference-overview.md)
- [<span data-ttu-id="4a76d-214">Trabalhe com intervalos usando a API JavaScript do Excel (fundamental)</span><span class="sxs-lookup"><span data-stu-id="4a76d-214">Work with ranges using the Excel JavaScript API (fundamental)</span></span>](excel-add-ins-ranges.md)
- [<span data-ttu-id="4a76d-215">Trabalhe com intervalos usando a API JavaScript do Excel (avançado)</span><span class="sxs-lookup"><span data-stu-id="4a76d-215">Work with ranges using the Excel JavaScript API (advanced)</span></span>](excel-add-ins-ranges-advanced.md)
