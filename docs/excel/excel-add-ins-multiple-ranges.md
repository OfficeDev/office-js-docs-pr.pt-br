---
title: Trabalhar simultaneamente com vários intervalos em suplementos do Excel
description: ''
ms.date: 09/04/2018
ms.openlocfilehash: f1217fc76d14269882a73ec5eb7758e519563456
ms.sourcegitcommit: 6870f0d96ed3da2da5a08652006c077a72d811b6
ms.translationtype: HT
ms.contentlocale: pt-BR
ms.lasthandoff: 12/21/2018
ms.locfileid: "27383222"
---
# <a name="work-with-multiple-ranges-simultaneously-in-excel-add-ins-preview"></a><span data-ttu-id="f2b5a-102">Trabalhar simultaneamente com vários intervalos em suplementos do Excel (Visualização)</span><span class="sxs-lookup"><span data-stu-id="f2b5a-102">Work with multiple ranges simultaneously in Excel add-ins (Preview)</span></span>

<span data-ttu-id="f2b5a-103">A biblioteca de JavaScript do Excel permite que o suplemento realize operações e defina propriedades, em vários intervalos simultaneamente.</span><span class="sxs-lookup"><span data-stu-id="f2b5a-103">The Excel JavaScript library enables your add-in to perform operations, and set properties, on multiple ranges simultaneously.</span></span> <span data-ttu-id="f2b5a-104">Os intervalos não precisam ser contíguos.</span><span class="sxs-lookup"><span data-stu-id="f2b5a-104">The ranges do not have to be contiguous.</span></span> <span data-ttu-id="f2b5a-105">Além de tornar seu código mais simples, essa maneira de definir uma propriedade é executada muito mais rapidamente do que definir a mesma propriedade individualmente para cada um dos intervalos.</span><span class="sxs-lookup"><span data-stu-id="f2b5a-105">In addition to making your code simpler, this way of setting a property runs much faster than setting the same property individually for each of the ranges.</span></span>

> [!NOTE]
> <span data-ttu-id="f2b5a-106">As APIs descritas neste artigo requerem a \*\* versão 1809 Build 10820.20000 clique para executar do Office 2016\*\* ou posterior.</span><span class="sxs-lookup"><span data-stu-id="f2b5a-106">The APIs described in this article require **Office 2016 Click-to-Run version 1809 Build 10820.20000** or later.</span></span> <span data-ttu-id="f2b5a-107">(Talvez seja necessário ingressar o [programa Office Insider](https://products.office.com/office-insider) para obter uma compilação apropriada.) Além disso, você deve carregar a versão beta da biblioteca JavaScript do Office [CDN Office.js](https://appsforoffice.microsoft.com/lib/beta/hosted/office.js).</span><span class="sxs-lookup"><span data-stu-id="f2b5a-107">(You may need to join the [Office Insider program](https://products.office.com/office-insider) to get an appropriate build.) Also, you must load the beta version of the Office JavaScript library from [Office.js CDN](https://appsforoffice.microsoft.com/lib/beta/hosted/office.js).</span></span> <span data-ttu-id="f2b5a-108">Por fim, ainda não temos páginas de referência para essas APIs.</span><span class="sxs-lookup"><span data-stu-id="f2b5a-108">Finally, we don't have reference pages for these APIs yet.</span></span> <span data-ttu-id="f2b5a-109">Mas o seguinte arquivo de tipo de definição tem descrições para eles: [beta office.d.ts](https://appsforoffice.microsoft.com/lib/beta/hosted/office.d.ts).</span><span class="sxs-lookup"><span data-stu-id="f2b5a-109">But the following definition type file has descriptions for them: [beta office.d.ts](https://appsforoffice.microsoft.com/lib/beta/hosted/office.d.ts).</span></span>

## <a name="rangeareas"></a><span data-ttu-id="f2b5a-110">RangeAreas</span><span class="sxs-lookup"><span data-stu-id="f2b5a-110">RangeAreas</span></span>

<span data-ttu-id="f2b5a-111">Um conjunto de intervalos (possivelmente não contíguos) é representado por um objeto `Excel.RangeAreas`.</span><span class="sxs-lookup"><span data-stu-id="f2b5a-111">A set of (possibly discontiguous) ranges is represented by an `Excel.RangeAreas` object.</span></span> <span data-ttu-id="f2b5a-112">Possui propriedades e métodos semelhantes ao tipo `Range` (muitos com os mesmos nomes ou semelhantes), mas foram feitos ajustes para:</span><span class="sxs-lookup"><span data-stu-id="f2b5a-112">It has properties and methods similar to the `Range` type (many with the same, or similar, names), but adjustments have been made to:</span></span>

- <span data-ttu-id="f2b5a-113">Os tipos de dados para propriedades e o comportamento dos setters e getters.</span><span class="sxs-lookup"><span data-stu-id="f2b5a-113">The data types for properties and the behavior of the setters and getters.</span></span>
- <span data-ttu-id="f2b5a-114">Os tipos de dados dos parâmetros do método e os comportamentos do método.</span><span class="sxs-lookup"><span data-stu-id="f2b5a-114">The data types of method parameters and the method behaviors.</span></span>
- <span data-ttu-id="f2b5a-115">Os tipos de dados de forma retornam valores.</span><span class="sxs-lookup"><span data-stu-id="f2b5a-115">The data types of method return values.</span></span>

<span data-ttu-id="f2b5a-116">Alguns exemplos:</span><span class="sxs-lookup"><span data-stu-id="f2b5a-116">Some examples:</span></span>

- <span data-ttu-id="f2b5a-117">`RangeAreas` tem uma propriedade `address` que retorna uma cadeia de caracteres delimitada por vírgula de intervalo de endereços, em vez de apenas um endereço como na propriedade`Range.address`.</span><span class="sxs-lookup"><span data-stu-id="f2b5a-117">`RangeAreas` has an `address` property that returns a comma-delimited string of range addresses, instead of just one address as with the `Range.address` property.</span></span>
- <span data-ttu-id="f2b5a-118">`RangeAreas` tem uma propriedade `dataValidation` que retorna um objeto `DataValidation` que representa a validação de dados de todos os intervalos em`RangeAreas`, se for consistente.</span><span class="sxs-lookup"><span data-stu-id="f2b5a-118">`RangeAreas` has a `dataValidation` property that returns a `DataValidation` object that represents the data validation of all the ranges in the `RangeAreas`, if it is consistent.</span></span> <span data-ttu-id="f2b5a-119">A propriedade é `null` se objetos idênticos `DataValidation` não forem aplicados a todos os intervalos em `RangeAreas`.</span><span class="sxs-lookup"><span data-stu-id="f2b5a-119">The property is `null` if identical `DataValidation` objects are not applied to all the all the ranges in the `RangeAreas`.</span></span> <span data-ttu-id="f2b5a-120">Esse é um princípio geral, mas não universal com o objeto `RangeAreas`: *se uma propriedade não têm valores consistentes em todos os todos os intervalos em `RangeAreas`, então será `null`.*</span><span class="sxs-lookup"><span data-stu-id="f2b5a-120">This is a general, but not universal, principle with the `RangeAreas` object: *If a property does not have consistent values on all the all the ranges in the `RangeAreas`, then it is `null`.*</span></span> <span data-ttu-id="f2b5a-121">Ver [ler as propriedades de RangeAreas](#read-properties-of-rangeareas) para mais informações e algumas exceções.</span><span class="sxs-lookup"><span data-stu-id="f2b5a-121">See [Read properties of RangeAreas](#read-properties-of-rangeareas) for more information and some exceptions.</span></span>
- <span data-ttu-id="f2b5a-122">`RangeAreas.cellCount` é o número total de células em todos os intervalos no `RangeAreas`.</span><span class="sxs-lookup"><span data-stu-id="f2b5a-122">`RangeAreas.cellCount` gets the total number of cells in all the ranges in the `RangeAreas`.</span></span>
- <span data-ttu-id="f2b5a-123">`RangeAreas.calculate` recalcula as células de todos os intervalos no `RangeAreas`.</span><span class="sxs-lookup"><span data-stu-id="f2b5a-123">`RangeAreas.calculate` recalculates the cells of all the ranges in the `RangeAreas`.</span></span>
- <span data-ttu-id="f2b5a-124">`RangeAreas.getEntireColumn` e `RangeAreas.getEntireRow` retornar outra `RangeAreas` objeto que representa todas as colunas (ou linhas) em todos os intervalos no `RangeAreas`.</span><span class="sxs-lookup"><span data-stu-id="f2b5a-124">`RangeAreas.getEntireColumn` and `RangeAreas.getEntireRow` return another `RangeAreas` object that represents all of the columns (or rows) in all the ranges in the `RangeAreas`.</span></span> <span data-ttu-id="f2b5a-125">Por exemplo, se `RangeAreas` representa "A1: C4" e "F14:L15" em seguida, `RangeAreas.getEntireColumn` retorna um objeto `RangeAreas` que representa "A:C" e "F:L".</span><span class="sxs-lookup"><span data-stu-id="f2b5a-125">For example, if the `RangeAreas` represents "A1:C4" and "F14:L15", then `RangeAreas.getEntireColumn` returns a `RangeAreas` object that represents "A:C" and "F:L".</span></span>
- <span data-ttu-id="f2b5a-126">`RangeAreas.copyFrom` pode ter o parâmetro `Range` ou `RangeAreas` que representam os intervalos de origem da operação de cópia.</span><span class="sxs-lookup"><span data-stu-id="f2b5a-126">`RangeAreas.copyFrom` can take either a `Range` or a `RangeAreas` parameter representing the source range(s) of the copy operation.</span></span>

#### <a name="complete-list-of-range-members-that-are-also-available-on-rangeareas"></a><span data-ttu-id="f2b5a-127">Lista completa de membros do intervalo que também estão disponíveis em RangeAreas</span><span class="sxs-lookup"><span data-stu-id="f2b5a-127">Complete list of Range members that are also available on RangeAreas</span></span>

##### <a name="properties"></a><span data-ttu-id="f2b5a-128">Propriedades</span><span class="sxs-lookup"><span data-stu-id="f2b5a-128">Properties</span></span>

<span data-ttu-id="f2b5a-129">Familiarize-se com as [Propriedades de leitura do RangeAreas](#read-properties-of-rangeareas) antes de escrever o código que lê as propriedades listadas.</span><span class="sxs-lookup"><span data-stu-id="f2b5a-129">Be familiar with [Read properties of RangeAreas](#read-properties-of-rangeareas) before you write code that reads any properties listed.</span></span> <span data-ttu-id="f2b5a-130">Existem sutilezas para o que é retornado.</span><span class="sxs-lookup"><span data-stu-id="f2b5a-130">There are subtleties to what gets returned.</span></span>

- <span data-ttu-id="f2b5a-131">address</span><span class="sxs-lookup"><span data-stu-id="f2b5a-131">address</span></span>
- <span data-ttu-id="f2b5a-132">addressLocal</span><span class="sxs-lookup"><span data-stu-id="f2b5a-132">addressLocal</span></span>
- <span data-ttu-id="f2b5a-133">cellCount</span><span class="sxs-lookup"><span data-stu-id="f2b5a-133">cellCount</span></span>
- <span data-ttu-id="f2b5a-134">conditionalFormats</span><span class="sxs-lookup"><span data-stu-id="f2b5a-134">conditionalFormats</span></span>
- <span data-ttu-id="f2b5a-135">context</span><span class="sxs-lookup"><span data-stu-id="f2b5a-135">context</span></span>
- <span data-ttu-id="f2b5a-136">dataValidation</span><span class="sxs-lookup"><span data-stu-id="f2b5a-136">dataValidation</span></span>
- <span data-ttu-id="f2b5a-137">formato</span><span class="sxs-lookup"><span data-stu-id="f2b5a-137">format</span></span>
- <span data-ttu-id="f2b5a-138">isEntireColumn</span><span class="sxs-lookup"><span data-stu-id="f2b5a-138">isEntireColumn</span></span>
- <span data-ttu-id="f2b5a-139">isEntireRow</span><span class="sxs-lookup"><span data-stu-id="f2b5a-139">isEntireRow</span></span>
- <span data-ttu-id="f2b5a-140">style</span><span class="sxs-lookup"><span data-stu-id="f2b5a-140">style</span></span>
- <span data-ttu-id="f2b5a-141">planilha</span><span class="sxs-lookup"><span data-stu-id="f2b5a-141">worksheet</span></span>

##### <a name="methods"></a><span data-ttu-id="f2b5a-142">Métodos</span><span class="sxs-lookup"><span data-stu-id="f2b5a-142">Methods</span></span>

<span data-ttu-id="f2b5a-143">Os métodos de intervalo na visualização são marcados.</span><span class="sxs-lookup"><span data-stu-id="f2b5a-143">Range methods in preview are marked.</span></span>

- <span data-ttu-id="f2b5a-144">calculate()</span><span class="sxs-lookup"><span data-stu-id="f2b5a-144">calculate()</span></span>
- <span data-ttu-id="f2b5a-145">clear()</span><span class="sxs-lookup"><span data-stu-id="f2b5a-145">clear()</span></span>
- <span data-ttu-id="f2b5a-146">convertDataTypeToText() (visualização)</span><span class="sxs-lookup"><span data-stu-id="f2b5a-146">convertDataTypeToText() (preview)</span></span>
- <span data-ttu-id="f2b5a-147">convertToLinkedDataType() (visualização)</span><span class="sxs-lookup"><span data-stu-id="f2b5a-147">convertToLinkedDataType() (preview)</span></span>
- <span data-ttu-id="f2b5a-148">copyFrom() (visualização)</span><span class="sxs-lookup"><span data-stu-id="f2b5a-148">copyFrom() (preview)</span></span>
- <span data-ttu-id="f2b5a-149">getEntireColumn()</span><span class="sxs-lookup"><span data-stu-id="f2b5a-149">getEntireColumn()</span></span>
- <span data-ttu-id="f2b5a-150">getEntireRow()</span><span class="sxs-lookup"><span data-stu-id="f2b5a-150">getEntireRow()</span></span>
- <span data-ttu-id="f2b5a-151">getIntersection()</span><span class="sxs-lookup"><span data-stu-id="f2b5a-151">getIntersection()</span></span>
- <span data-ttu-id="f2b5a-152">getIntersectionOrNullObject()</span><span class="sxs-lookup"><span data-stu-id="f2b5a-152">getIntersectionOrNullObject()</span></span>
- <span data-ttu-id="f2b5a-153">getOffsetRange() (chamada getOffsetRangeAreas no objeto RangeAreas)</span><span class="sxs-lookup"><span data-stu-id="f2b5a-153">getOffsetRange() (named getOffsetRangeAreas on the RangeAreas object)</span></span>
- <span data-ttu-id="f2b5a-154">getSpecialCells() (visualização)</span><span class="sxs-lookup"><span data-stu-id="f2b5a-154">getSpecialCells() (preview)</span></span>
- <span data-ttu-id="f2b5a-155">getSpecialCellsOrNullObject() (visualização)</span><span class="sxs-lookup"><span data-stu-id="f2b5a-155">getSpecialCellsOrNullObject() (preview)</span></span>
- <span data-ttu-id="f2b5a-156">getTables() (visualização)</span><span class="sxs-lookup"><span data-stu-id="f2b5a-156">getTables() (preview)</span></span>
- <span data-ttu-id="f2b5a-157">getUsedRange() (chamada getUsedRangeAreas no objeto RangeAreas)</span><span class="sxs-lookup"><span data-stu-id="f2b5a-157">getUsedRange() (named getUsedRangeAreas on the RangeAreas object)</span></span>
- <span data-ttu-id="f2b5a-158">getUsedRangeOrNullObject() (chamada getUsedRangeAreasOrNullObject no objeto RangeAreas)</span><span class="sxs-lookup"><span data-stu-id="f2b5a-158">getUsedRangeOrNullObject() (named getUsedRangeAreasOrNullObject on the RangeAreas object)</span></span>
- <span data-ttu-id="f2b5a-159">load()</span><span class="sxs-lookup"><span data-stu-id="f2b5a-159">load()</span></span>
- <span data-ttu-id="f2b5a-160">set()</span><span class="sxs-lookup"><span data-stu-id="f2b5a-160">set()</span></span>
- <span data-ttu-id="f2b5a-161">setDirty() (visualização)</span><span class="sxs-lookup"><span data-stu-id="f2b5a-161">setDirty() (preview)</span></span>
- <span data-ttu-id="f2b5a-162">toJSON()</span><span class="sxs-lookup"><span data-stu-id="f2b5a-162">toJSON()</span></span>
- <span data-ttu-id="f2b5a-163">track()</span><span class="sxs-lookup"><span data-stu-id="f2b5a-163">track()</span></span>
- <span data-ttu-id="f2b5a-164">untrack()</span><span class="sxs-lookup"><span data-stu-id="f2b5a-164">untrack()</span></span>

### <a name="rangearea-specific-properties-and-methods"></a><span data-ttu-id="f2b5a-165">Métodos e propriedades específicos do RangeArea</span><span class="sxs-lookup"><span data-stu-id="f2b5a-165">RangeArea-specific properties and methods</span></span>

<span data-ttu-id="f2b5a-166">O tipo `RangeAreas` tem alguns métodos e propriedades que não estão no objeto `Range`.</span><span class="sxs-lookup"><span data-stu-id="f2b5a-166">The `RangeAreas` type has some properties and methods that are not on the `Range` object.</span></span> <span data-ttu-id="f2b5a-167">Esta é a seleção deles:</span><span class="sxs-lookup"><span data-stu-id="f2b5a-167">The following is a selection of them:</span></span>

- <span data-ttu-id="f2b5a-168">`areas`: O objeto `RangeCollection` que contém todos os intervalos representados pelo objeto `RangeAreas`.</span><span class="sxs-lookup"><span data-stu-id="f2b5a-168">`areas`: A `RangeCollection` object that contains all of the ranges represented by the `RangeAreas` object.</span></span> <span data-ttu-id="f2b5a-169">O objeto `RangeCollection` também é novidade e é semelhante a outros objetos do conjunto do Excel.</span><span class="sxs-lookup"><span data-stu-id="f2b5a-169">The `RangeCollection` object is also new and is similar to other Excel collection objects.</span></span> <span data-ttu-id="f2b5a-170">É uma propriedade `items` que é uma matriz de objetos `Range` que representam os intervalos.</span><span class="sxs-lookup"><span data-stu-id="f2b5a-170">It has an `items` property which is an array of `Range` objects representing the ranges.</span></span>
- <span data-ttu-id="f2b5a-171">`areaCount`: O número total de intervalos em `RangeAreas`.</span><span class="sxs-lookup"><span data-stu-id="f2b5a-171">`areaCount`: The total number of ranges in the `RangeAreas`.</span></span>
- <span data-ttu-id="f2b5a-172">`getOffsetRangeAreas`: Funciona como [Range.getOffsetRange](/javascript/api/excel/excel.range#getoffsetrange-rowoffset--columnoffset-), exceto pelo fato de que o `RangeAreas` é retornado e contém os intervalos que são todos os deslocamentos de um dos intervalos do `RangeAreas` original.</span><span class="sxs-lookup"><span data-stu-id="f2b5a-172">`getOffsetRangeAreas`: Works just like [Range.getOffsetRange](/javascript/api/excel/excel.range#getoffsetrange-rowoffset--columnoffset-), except that a `RangeAreas` is returned and it contains ranges that are each offset from one of the ranges in the original `RangeAreas`.</span></span>

## <a name="create-rangeareas-and-set-properties"></a><span data-ttu-id="f2b5a-173">Criar RangeAreas e definir propriedades</span><span class="sxs-lookup"><span data-stu-id="f2b5a-173">Create RangeAreas and set properties</span></span>

<span data-ttu-id="f2b5a-174">Você pode criar o objeto`RangeAreas` de duas maneiras básicas:</span><span class="sxs-lookup"><span data-stu-id="f2b5a-174">You can create `RangeAreas` object in two basic ways:</span></span>

- <span data-ttu-id="f2b5a-175">Ligue `Worksheet.getRanges()` e encaminhe-o em uma cadeia de caracteres com endereços de intervalo separado por vírgula.</span><span class="sxs-lookup"><span data-stu-id="f2b5a-175">Call `Worksheet.getRanges()` and pass it a string with comma-delimited range addresses.</span></span> <span data-ttu-id="f2b5a-176">Se algum intervalo que você deseja incluir tiver sido feito em um [NamedItem](https://docs.microsoft.com/javascript/api/excel/excel.nameditem), você poderá incluir o nome, em vez do endereço, cadeia de caracteres.</span><span class="sxs-lookup"><span data-stu-id="f2b5a-176">If any range you want to include has been made into a [NamedItem](https://docs.microsoft.com/javascript/api/excel/excel.nameditem), you can include the name, instead of the address, in the string.</span></span>
- <span data-ttu-id="f2b5a-177">Chamar `Workbook.getSelectedRanges()`.</span><span class="sxs-lookup"><span data-stu-id="f2b5a-177">Call `Workbook.getSelectedRanges()`.</span></span> <span data-ttu-id="f2b5a-178">Esse método retornará um `RangeAreas` representando todos os intervalos selecionados na planilha ativa no momento.</span><span class="sxs-lookup"><span data-stu-id="f2b5a-178">This method returns a `RangeAreas` representing all the ranges that are selected on the currently active worksheet.</span></span>

<span data-ttu-id="f2b5a-179">Quando você tiver um objeto `RangeAreas`, você pode criar outros usando os métodos de objeto que retornam `RangeAreas` como `getOffsetRangeAreas` e `getIntersection`.</span><span class="sxs-lookup"><span data-stu-id="f2b5a-179">Once you have a `RangeAreas` object, you can create others using the methods on the object that return `RangeAreas` such as `getOffsetRangeAreas` and `getIntersection`.</span></span>

> [!NOTE]
> <span data-ttu-id="f2b5a-180">É possível adicionar diretamente intervalos adicionais para um objeto `RangeAreas`.</span><span class="sxs-lookup"><span data-stu-id="f2b5a-180">You cannot directly add additional ranges to a `RangeAreas` object.</span></span> <span data-ttu-id="f2b5a-181">Por exemplo, o conjunto `RangeAreas.areas` não tem um método`add`.</span><span class="sxs-lookup"><span data-stu-id="f2b5a-181">For example, the collection in `RangeAreas.areas` does not have an `add` method.</span></span>

> [!WARNING]
> <span data-ttu-id="f2b5a-182">Tente adicionar ou excluir membros diretamente à matriz`RangeAreas.areas.items`.</span><span class="sxs-lookup"><span data-stu-id="f2b5a-182">Do not attempt to directly add or delete members of the the `RangeAreas.areas.items` array.</span></span> <span data-ttu-id="f2b5a-183">Isso levará a um comportamento indesejável no seu código.</span><span class="sxs-lookup"><span data-stu-id="f2b5a-183">This will lead to undesirable behavior in your code.</span></span> <span data-ttu-id="f2b5a-184">Por exemplo, é possível enviar um objeto adicional `Range` para a matriz, mas isso causará erros porque as propriedades e métodos `RangeAreas` se comportam como se o novo item não estivesse ali.</span><span class="sxs-lookup"><span data-stu-id="f2b5a-184">For example, it is possible to push an additional `Range` object onto the array, but doing so will cause errors because `RangeAreas` properties and methods behave as if the new item isn't there.</span></span> <span data-ttu-id="f2b5a-185">Por exemplo, a propriedade `areaCount` não inclui intervalos transferidos dessa maneira e o `RangeAreas.getItemAt(index)` gera um erro se `index` for maior que `areasCount-1`.</span><span class="sxs-lookup"><span data-stu-id="f2b5a-185">For example, the `areaCount` property does not include ranges pushed in this way, and the `RangeAreas.getItemAt(index)` throws an error if `index` is larger than `areasCount-1`.</span></span> <span data-ttu-id="f2b5a-186">Da mesma forma, excluir um objeto `Range` na matriz `RangeAreas.areas.items` obtendo uma referência a ele e chamando seu método `Range.delete` causa bugs: embora o `Range`objeto\* seja \*excluído, as propriedades e métodos do objeto pai `RangeAreas` se comportam ou tentam se comportar, como se ele ainda existisse.</span><span class="sxs-lookup"><span data-stu-id="f2b5a-186">Similarly, deleting a `Range` object in the `RangeAreas.areas.items` array by getting a reference to it and calling its `Range.delete` method causes bugs: although the `Range` object *is* deleted, the properties and methods of the parent `RangeAreas` object behave, or try to, as if it is still in existence.</span></span> <span data-ttu-id="f2b5a-187">Por exemplo, se o seu código chamar `RangeAreas.calculate`, o Office tentará calcular o intervalo, mas haverá erro porque o objeto de intervalo desapareceu.</span><span class="sxs-lookup"><span data-stu-id="f2b5a-187">For example, if your code calls `RangeAreas.calculate`, Office will try to calculate the range, but will error because the range object is gone.</span></span>

<span data-ttu-id="f2b5a-188">Definir uma propriedade em um `RangeAreas` define a propriedade correspondente em todos os intervalos no conjunto `RangeAreas.areas`.</span><span class="sxs-lookup"><span data-stu-id="f2b5a-188">Setting a property on a `RangeAreas` sets the corresponding property on all the ranges in the `RangeAreas.areas` collection.</span></span>

<span data-ttu-id="f2b5a-189">A seguir, um exemplo de configuração de uma propriedade em vários intervalos.</span><span class="sxs-lookup"><span data-stu-id="f2b5a-189">The following is an example of setting a property on multiple ranges.</span></span> <span data-ttu-id="f2b5a-190">A função realça os intervalos **F3:F5** e **H3:H5**.</span><span class="sxs-lookup"><span data-stu-id="f2b5a-190">The function highlights the ranges **F3:F5** and **H3:H5**.</span></span>

```js
Excel.run(function (context) {
    var sheet = context.workbook.worksheets.getActiveWorksheet();
    var rangeAreas = sheet.getRanges("F3:F5, H3:H5");
    rangeAreas.format.fill.color = "pink";

    return context.sync();
})
```

<span data-ttu-id="f2b5a-191">Este exemplo se aplica a cenários nos quais você pode codificar os endereços de intervalo para os quais você passa para `getRanges` ou facilmente calculá-los no tempo de execução.</span><span class="sxs-lookup"><span data-stu-id="f2b5a-191">This example applies to scenarios in which you can hard code the range addresses that you pass to `getRanges` or easily calculate them at runtime.</span></span> <span data-ttu-id="f2b5a-192">Alguns dos cenários em que isso pode ser verdadeiro incluem:</span><span class="sxs-lookup"><span data-stu-id="f2b5a-192">Some of the scenarios in which this would be true include:</span></span> 

- <span data-ttu-id="f2b5a-193">O código é executado no contexto de um modelo conhecido.</span><span class="sxs-lookup"><span data-stu-id="f2b5a-193">The code runs in the context of a known template.</span></span>
- <span data-ttu-id="f2b5a-194">O código é executado no contexto de dados importados, em que o esquema dos dados é conhecido.</span><span class="sxs-lookup"><span data-stu-id="f2b5a-194">The code runs in the context of imported data where the schema of the data is known.</span></span>

<span data-ttu-id="f2b5a-195">Quando você não pode saber no tempo de codificação quais intervalos você precisa operar, você deve descobri-los em tempo de execução.</span><span class="sxs-lookup"><span data-stu-id="f2b5a-195">When you can't know at coding-time which ranges you need to operate on, you must discover them at runtime.</span></span> <span data-ttu-id="f2b5a-196">A seção a seguir descreve esses cenários.</span><span class="sxs-lookup"><span data-stu-id="f2b5a-196">The next section discusses these scenarios.</span></span>

### <a name="discover-range-areas-programmatically"></a><span data-ttu-id="f2b5a-197">Descubra as áreas de intervalo por programação</span><span class="sxs-lookup"><span data-stu-id="f2b5a-197">Discover range areas programmatically</span></span>

<span data-ttu-id="f2b5a-198">Os métodos `Range.getSpecialCells()` e `Range.getSpecialCellsOrNullObject()` permitem localizar no tempo de execução os intervalos nos quais você deseja operar, com base nas características das células e no tipo de valores das células.</span><span class="sxs-lookup"><span data-stu-id="f2b5a-198">The `Range.getSpecialCells()` and `Range.getSpecialCellsOrNullObject()` methods enable you to find at runtime the ranges that you want to operate on based on the characteristics of the cells and the type of the values of the cells.</span></span> <span data-ttu-id="f2b5a-199">Aqui estão as assinaturas dos métodos do arquivo de tipos de dados TypeScript:</span><span class="sxs-lookup"><span data-stu-id="f2b5a-199">Here are the signatures of the methods from the TypeScript data types file:</span></span>

```typescript
getSpecialCells(cellType: Excel.SpecialCellType, cellValueType?: Excel.SpecialCellValueType): Excel.RangeAreas;
```

```typescript
getSpecialCellsOrNullObject(cellType: Excel.SpecialCellType, cellValueType?: Excel.SpecialCellValueType): Excel.RangeAreas;
```

<span data-ttu-id="f2b5a-200">Este é um exemplo de como usar a primeira.</span><span class="sxs-lookup"><span data-stu-id="f2b5a-200">The following is an example of using the first one.</span></span> <span data-ttu-id="f2b5a-201">Sobre este código, observe:</span><span class="sxs-lookup"><span data-stu-id="f2b5a-201">About this code, note:</span></span>

- <span data-ttu-id="f2b5a-202">Ele limita a parte da planilha que precisa ser pesquisada chamando primeiro `Worksheet.getUsedRange` e chamando `getSpecialCells` para apenas esse intervalo.</span><span class="sxs-lookup"><span data-stu-id="f2b5a-202">It limits the part of the sheet that needs to be searched by first calling `Worksheet.getUsedRange` and calling `getSpecialCells` for only that range.</span></span>
- <span data-ttu-id="f2b5a-203">Ele passa como um parâmetro para a versão `getSpecialCells` de seqüência de caracteres de um valor do enum `Excel.SpecialCellType`.</span><span class="sxs-lookup"><span data-stu-id="f2b5a-203">It passes as a parameter to `getSpecialCells` the string version of a value from the `Excel.SpecialCellType` enum.</span></span> <span data-ttu-id="f2b5a-204">Alguns dos outros valores que podem ser passados ​​são "Blanks" para células vazias, "Constantes" para células com valores literais em vez de fórmulas e "SameConditionalFormat" para células que possuem a mesma formatação condicional que a primeira célula em `usedRange`.</span><span class="sxs-lookup"><span data-stu-id="f2b5a-204">Some of the other values that could be passed instead are "Blanks" for empty cells, "Constants" for cells with literal values instead of formulas, and "SameConditionalFormat" for cells that have the same conditional formatting as the first cell in the `usedRange`.</span></span> <span data-ttu-id="f2b5a-205">A primeira célula é a célula superior esquerda.</span><span class="sxs-lookup"><span data-stu-id="f2b5a-205">The first cell is the upper leftmost cell.</span></span> <span data-ttu-id="f2b5a-206">Para uma lista completa dos valores na enumeração, confira [beta office.d.ts](https://appsforoffice.microsoft.com/lib/beta/hosted/office.d.ts).</span><span class="sxs-lookup"><span data-stu-id="f2b5a-206">For a complete list of the values in the enum, see [beta office.d.ts](https://appsforoffice.microsoft.com/lib/beta/hosted/office.d.ts).</span></span>
- <span data-ttu-id="f2b5a-207">O método `getSpecialCells` retorna um objeto `RangeAreas`, então todas as células com fórmulas serão coloridas de rosa, mesmo que não sejam todas contíguas.</span><span class="sxs-lookup"><span data-stu-id="f2b5a-207">The `getSpecialCells` method returns a `RangeAreas` object, so all of the cells with formulas will be colored pink even if they are not all contiguous.</span></span> 

```js
Excel.run(function (context) {
    var sheet = context.workbook.worksheets.getActiveWorksheet();
    var usedRange = sheet.getUsedRange();
    var formulaRanges = usedRange.getSpecialCells("Formulas");
    formulaRanges.format.fill.color = "pink";

    return context.sync();
})
```

<span data-ttu-id="f2b5a-208">Às vezes, o intervalo não possui *nenhuma* célula com a característica desejada.</span><span class="sxs-lookup"><span data-stu-id="f2b5a-208">Sometimes the range doesn't have *any* cells with the targeted characteristic.</span></span> <span data-ttu-id="f2b5a-209">Se `getSpecialCells` não encontrar nenhuma, ele lançará um erro **ItemNotFound**.</span><span class="sxs-lookup"><span data-stu-id="f2b5a-209">If `getSpecialCells` doesn't find any, it throws an **ItemNotFound** error.</span></span> <span data-ttu-id="f2b5a-210">Isso iria desviar o fluxo de controle para um bloco / método `catch`, se houver um.</span><span class="sxs-lookup"><span data-stu-id="f2b5a-210">This would divert the flow of control to a `catch` block/method, if there is one.</span></span> <span data-ttu-id="f2b5a-211">Se não houver, o erro interrompe a função.</span><span class="sxs-lookup"><span data-stu-id="f2b5a-211">If there isn't, the error halts the function.</span></span> <span data-ttu-id="f2b5a-212">Pode haver cenários em que lançar o erro é exatamente o que você quer que aconteça quando não houver células com a característica de destino.</span><span class="sxs-lookup"><span data-stu-id="f2b5a-212">There may be scenarios in which throwing the error is exactly what you want to happen when there are no cells with the targeted characteristic.</span></span> 

<span data-ttu-id="f2b5a-213">Mas em cenários em que é normal, mas talvez incomum, não haver células correspondentes; seu código deve verificar essa possibilidade e lidar com isso sem causar erro.</span><span class="sxs-lookup"><span data-stu-id="f2b5a-213">But in scenarios in which it is normal, but perhaps uncommon, for there to be no matching cells; your code should check for this possibility and handle it gracefully without throwing an error.</span></span> <span data-ttu-id="f2b5a-214">Para essas situações, use o método `getSpecialCellsOrNullObject` e teste a propriedade`RangeAreas.isNullObject`.</span><span class="sxs-lookup"><span data-stu-id="f2b5a-214">For these scenarios, use the `getSpecialCellsOrNullObject` method and test the `RangeAreas.isNullObject` property.</span></span> <span data-ttu-id="f2b5a-215">Apresentamos um exemplo a seguir.</span><span class="sxs-lookup"><span data-stu-id="f2b5a-215">The following is an example.</span></span> <span data-ttu-id="f2b5a-216">Observação sobre o código:</span><span class="sxs-lookup"><span data-stu-id="f2b5a-216">Note about this code:</span></span>

- <span data-ttu-id="f2b5a-217">O método `getSpecialCellsOrNullObject` sempre retorna um objeto de proxy, portanto, `null` nunca está no sentido comum do JavaScript.</span><span class="sxs-lookup"><span data-stu-id="f2b5a-217">The `getSpecialCellsOrNullObject` method always returns a proxy object, so it is never `null` in the ordinary JavaScript sense.</span></span> <span data-ttu-id="f2b5a-218">Mas se nenhuma célula de correspondência for encontrada, as propriedades do objeto`isNullObject` serão definida como `true`.</span><span class="sxs-lookup"><span data-stu-id="f2b5a-218">But if no matching cells are found, the `isNullObject` property of the object is set to `true`.</span></span>
- <span data-ttu-id="f2b5a-219">Ele chama `context.sync` *antes* de testar a propriedade`isNullObject`.</span><span class="sxs-lookup"><span data-stu-id="f2b5a-219">It calls `context.sync` *before* it tests the `isNullObject` property.</span></span> <span data-ttu-id="f2b5a-220">Esse é um requisito com todos os métodos e propriedades `*OrNullObject`, pois sempre terá que carregar e sincronizar as propriedades na ordem para lê-la.</span><span class="sxs-lookup"><span data-stu-id="f2b5a-220">This is a requirement with all `*OrNullObject` methods and properties, because you always have to load and sync a property in order to read it.</span></span> <span data-ttu-id="f2b5a-221">No entanto, não é necessário carregar *explicitamente* a propriedade`isNullObject`.</span><span class="sxs-lookup"><span data-stu-id="f2b5a-221">However, it is not necessary to *explicitly* load the `isNullObject` property.</span></span> <span data-ttu-id="f2b5a-222">Será carregado automaticamente pelo `context.sync` mesmo se `load` não for chamado no objeto.</span><span class="sxs-lookup"><span data-stu-id="f2b5a-222">It is automatically loaded by the `context.sync` even if `load` is not called on the object.</span></span> <span data-ttu-id="f2b5a-223">Para saber mais, confira [ \*OrNullObject](https://docs.microsoft.com/office/dev/add-ins/excel/excel-add-ins-advanced-concepts#42ornullobject-methods).</span><span class="sxs-lookup"><span data-stu-id="f2b5a-223">For more information, see [\*OrNullObject](https://docs.microsoft.com/office/dev/add-ins/excel/excel-add-ins-advanced-concepts#42ornullobject-methods).</span></span>
- <span data-ttu-id="f2b5a-224">Você pode testar esse código selecionando primeiro um intervalo sem células de fórmula e executando-o.</span><span class="sxs-lookup"><span data-stu-id="f2b5a-224">You can test this code by first selecting a range that has no formula cells and running it.</span></span> <span data-ttu-id="f2b5a-225">Selecione um intervalo que tem pelo menos uma célula com uma fórmula e execute novamente.</span><span class="sxs-lookup"><span data-stu-id="f2b5a-225">Then select a range that has at least one cell with a formula and run it again.</span></span>

```js
Excel.run(function (context) {
    var range = context.workbook.getSelectedRange();
    var formulaRanges = range.getSpecialCellsOrNullObject("Formulas");
    return context.sync()
        .then(function() {
            if (formulaRanges.isNullObject) {
                console.log("No cells have formulas");
            }
            else {
                formulaRanges.format.fill.color = "pink";
            }
        })
        .then(context.sync);
})
```

<span data-ttu-id="f2b5a-226">Para manter a simplicidade, todos os outros exemplos deste artigo usam o método `getSpecialCells` em vez de `getSpecialCellsOrNullObject`.</span><span class="sxs-lookup"><span data-stu-id="f2b5a-226">For simplicity, all other examples in this article use the `getSpecialCells` method instead of  `getSpecialCellsOrNullObject`.</span></span>

#### <a name="narrow-the-target-cells-with-cell-value-types"></a><span data-ttu-id="f2b5a-227">Restringir as células de destino com tipos de valor de célula</span><span class="sxs-lookup"><span data-stu-id="f2b5a-227">Narrow the target cells with cell value types</span></span>

<span data-ttu-id="f2b5a-228">Há um segundo parâmetro opcional, do tipo de enumeração  `Excel.SpecialCellValueType`, que restringe ainda mais as células de destino.</span><span class="sxs-lookup"><span data-stu-id="f2b5a-228">There is an optional second parameter, of enum type `Excel.SpecialCellValueType`, that further narrows down the cells to target.</span></span> <span data-ttu-id="f2b5a-229">Você pode usá-lo somente quando você passar por  "Fórmulas" ou "Constantes" para `getSpecialCells` ou `getSpecialCellsOrNullObject`.</span><span class="sxs-lookup"><span data-stu-id="f2b5a-229">You can use it only when you pass either "Formulas" or "Constants" to `getSpecialCells` or `getSpecialCellsOrNullObject`.</span></span> <span data-ttu-id="f2b5a-230">O parâmetro especifica que você deseja apenas células com determinados tipos de valores.</span><span class="sxs-lookup"><span data-stu-id="f2b5a-230">The parameter specifies that you only want cells with certain types of values.</span></span> <span data-ttu-id="f2b5a-231">Há quatro tipos básicos: "Erro", "Lógica" (ou seja, booliano), "Números" e "Texto".</span><span class="sxs-lookup"><span data-stu-id="f2b5a-231">There are four basic types: "Error", "Logical" (which means boolean), "Numbers", and "Text".</span></span> <span data-ttu-id="f2b5a-232">(O enum tem outros valores além desses quatro que são discutidos abaixo.) O seguinte é um exemplo.</span><span class="sxs-lookup"><span data-stu-id="f2b5a-232">(The enum has other values besides these four which are discussed below.) The following is an example.</span></span> <span data-ttu-id="f2b5a-233">Sobre este código, observe:</span><span class="sxs-lookup"><span data-stu-id="f2b5a-233">About this code, note:</span></span>

- <span data-ttu-id="f2b5a-234">Ele apenas irá realçar células que contêm um valor numérico literal.</span><span class="sxs-lookup"><span data-stu-id="f2b5a-234">It will only highlight cells that have a literal number value.</span></span> <span data-ttu-id="f2b5a-235">Ele não destacará as células que têm uma fórmula (mesmo se o resultado for um número) ou células de estado booliano, de texto ou de erro.</span><span class="sxs-lookup"><span data-stu-id="f2b5a-235">It will not highlight cells that have a formula (even if the result is a number) or a boolean, text, or error state cells.</span></span>
- <span data-ttu-id="f2b5a-236">Para testar o código, certifique-se de que a planilha tenha algumas células com valores numéricos literais, algumas com outros tipos de valores literais e algumas com fórmulas.</span><span class="sxs-lookup"><span data-stu-id="f2b5a-236">To test the code, be sure the worksheet has some cells with literal number values, some with other kinds of literal values, and some with formulas.</span></span>

```js
Excel.run(function (context) {
    var sheet = context.workbook.worksheets.getActiveWorksheet();
    var usedRange = sheet.getUsedRange();
    var constantNumberRanges = usedRange.getSpecialCells("Constants", "Numbers");
    constantNumberRanges.format.fill.color = "pink";

    return context.sync();
})
```

<span data-ttu-id="f2b5a-237">Às vezes, você precisa operar com mais de um tipo de valor de célula, como todas as células com valor de texto e com valor booliano ("Lógico").</span><span class="sxs-lookup"><span data-stu-id="f2b5a-237">Sometimes you need to operate on more than one cell value type, such as all text-valued and all boolean-valued ("Logical") cells.</span></span> <span data-ttu-id="f2b5a-238">A enumeração `Excel.SpecialCellValueType` tem valores que permitem que você combine tipos.</span><span class="sxs-lookup"><span data-stu-id="f2b5a-238">The `Excel.SpecialCellValueType` enum has values that let you combine types.</span></span> <span data-ttu-id="f2b5a-239">Por exemplo, "LogicalText" segmentará todas as células booleanas e todas com valor de texto.</span><span class="sxs-lookup"><span data-stu-id="f2b5a-239">For example, "LogicalText" will target all boolean and all text-valued cells.</span></span> <span data-ttu-id="f2b5a-240">Você pode combinar dois ou três dos quatro tipos básicos.</span><span class="sxs-lookup"><span data-stu-id="f2b5a-240">You can combine any two or any three of the four basic types.</span></span> <span data-ttu-id="f2b5a-241">Os nomes desses valores de enumeração que combinam tipos básicos estão sempre em ordem alfabética.</span><span class="sxs-lookup"><span data-stu-id="f2b5a-241">The names of these enum values that combine basic types are always in alphabetical order.</span></span> <span data-ttu-id="f2b5a-242">Portanto, para combinar células com valor de erro, com valor de texto e valores boolianos, use "ErrorLogicalText", não "LogicalErrorText" ou "TextErrorLogical".</span><span class="sxs-lookup"><span data-stu-id="f2b5a-242">So to combine error-valued, text-valued, and boolean-valued cells, use "ErrorLogicalText", not "LogicalErrorText" or "TextErrorLogical".</span></span> <span data-ttu-id="f2b5a-243">O parâmetro padrão de "Todos" combina todos os quatro tipos.</span><span class="sxs-lookup"><span data-stu-id="f2b5a-243">The default parameter of "All" combines all four types.</span></span> <span data-ttu-id="f2b5a-244">O exemplo a seguir destaca todas as células com fórmulas que produzem valores ou números boolianos:</span><span class="sxs-lookup"><span data-stu-id="f2b5a-244">The following example highlights all cells with formulas that produce number or boolean values:</span></span>

```js
Excel.run(function (context) {
    var sheet = context.workbook.worksheets.getActiveWorksheet();
    var usedRange = sheet.getUsedRange();
    var formulaLogicalNumberRanges = usedRange.getSpecialCells("Formulas", "LogicalNumbers");
    formulaLogicalNumberRanges.format.fill.color = "pink";

    return context.sync();
})
```

> [!NOTE]
> <span data-ttu-id="f2b5a-245">O parâmetro `Excel.SpecialCellValueType` só poderá ser usado se o parâmetro `Excel.SpecialCellType` for “Fórmulas” ou “Constantes”</span><span class="sxs-lookup"><span data-stu-id="f2b5a-245">The `Excel.SpecialCellValueType` parameter can only be used if the `Excel.SpecialCellType` parameter is "Formulas" or "Constants".</span></span>

### <a name="get-rangeareas-within-rangeareas"></a><span data-ttu-id="f2b5a-246">Obter RangeAreas dentro de RangeAreas</span><span class="sxs-lookup"><span data-stu-id="f2b5a-246">Get RangeAreas within RangeAreas</span></span>

<span data-ttu-id="f2b5a-247">O tipo `RangeAreas` em si também possui métodos `getSpecialCells` e `getSpecialCellsOrNullObject` que usam os mesmos dois parâmetros.</span><span class="sxs-lookup"><span data-stu-id="f2b5a-247">The `RangeAreas` type itself also has `getSpecialCells` and `getSpecialCellsOrNullObject` methods which take the same two parameters.</span></span> <span data-ttu-id="f2b5a-248">Esses métodos retornam todas as células de destino de todos os intervalos no conjunto`RangeAreas.areas`.</span><span class="sxs-lookup"><span data-stu-id="f2b5a-248">These methods return all the targeted cells from all of the ranges in the `RangeAreas.areas` collection.</span></span> <span data-ttu-id="f2b5a-249">Há uma pequena diferença no comportamento dos métodos quando chamado em um objeto `RangeAreas` em vez de um objeto`Range`: quando você passa "SameConditionalFormat" como o primeiro parâmetro, o método retorna todas as células que têm a mesma formatação condicional que a célula superior esquerda \* no primeiro intervalo na `RangeAreas.areas`coleção\*.</span><span class="sxs-lookup"><span data-stu-id="f2b5a-249">There is one small difference in the behavior of the methods when called on a `RangeAreas` object instead of a `Range` object: when you pass "SameConditionalFormat" as the first parameter, the method returns all cells that have the same conditional formatting as the upper leftmost cell *in the first range in the `RangeAreas.areas` collection*.</span></span> <span data-ttu-id="f2b5a-250">O mesmo ponto se aplica a "SameDataValidation": quando passado para `Range.getSpecialCells`, ele retorna todas as células que possuem a mesma regra de validação de dados que a célula superior esquerda \* do intervalo\*.</span><span class="sxs-lookup"><span data-stu-id="f2b5a-250">The same point applies to "SameDataValidation": when passed to `Range.getSpecialCells`, it returns all cells that have the same data validation rule as the upper leftmost cell *in the range*.</span></span> <span data-ttu-id="f2b5a-251">Mas quando é passado para `RangeAreas.getSpecialCells`, retorna todas as células que têm a mesma regra de validação de dados que a célula superior esquerda \*no primeiro intervalo do conjunto`RangeAreas.areas` \*.</span><span class="sxs-lookup"><span data-stu-id="f2b5a-251">But when it is passed to `RangeAreas.getSpecialCells`, it returns all cells that have the same data validation rule as the upper leftmost cell *in the first range in the `RangeAreas.areas` collection*.</span></span>

## <a name="read-properties-of-rangeareas"></a><span data-ttu-id="f2b5a-252">Ler propriedades de RangeAreas</span><span class="sxs-lookup"><span data-stu-id="f2b5a-252">Read properties of RangeAreas</span></span>

<span data-ttu-id="f2b5a-253">A leitura de valores de propriedade `RangeAreas` requer cuidados, porque uma determinada propriedade pode ter valores diferentes para intervalos diferentes dentro de`RangeAreas`.</span><span class="sxs-lookup"><span data-stu-id="f2b5a-253">Reading property values of `RangeAreas` requires care, because a given property may have different values for different ranges within the `RangeAreas`.</span></span> <span data-ttu-id="f2b5a-254">A regra geral é que, se um valor consistente *puder* ser retornado, ele será retornado.</span><span class="sxs-lookup"><span data-stu-id="f2b5a-254">The general rule is that if a consistent value *can* be returned it will be returned.</span></span> <span data-ttu-id="f2b5a-255">Por exemplo, no código a seguir, O código RGB para pink (`#FFC0CB`) e `true` será registrado no console porque ambos os intervalos no objeto `RangeAreas` têm um preenchimento rosa e ambos são colunas inteiras.</span><span class="sxs-lookup"><span data-stu-id="f2b5a-255">For example, in the following code, The RGB code for pink (`#FFC0CB`) and `true` will be logged to the console because both the ranges in the `RangeAreas` object have a pink fill and both are entire columns.</span></span>

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

<span data-ttu-id="f2b5a-256">As coisas ficam mais complicadas quando a consistência não é possível.</span><span class="sxs-lookup"><span data-stu-id="f2b5a-256">Things get more complicated when consistency isn't possible.</span></span> <span data-ttu-id="f2b5a-257">O comportamento das propriedades `RangeAreas` seguem estes três princípios de três:</span><span class="sxs-lookup"><span data-stu-id="f2b5a-257">The behavior of `RangeAreas` properties follows these three principles:</span></span>

- <span data-ttu-id="f2b5a-258">Uma propriedade booliana de um `RangeAreas`retorno de objeto `false`, a menos que a propriedade seja verdadeira para todos os intervalos de membro.</span><span class="sxs-lookup"><span data-stu-id="f2b5a-258">A boolean property of a `RangeAreas` object returns `false` unless the property is true for all the member ranges.</span></span>
- <span data-ttu-id="f2b5a-259">Propriedades não boolianas, com exceção da propriedade `address`, retornam `null`, a menos que a propriedade correspondente em todos os intervalos de membros tenha o mesmo valor.</span><span class="sxs-lookup"><span data-stu-id="f2b5a-259">Non-boolean properties, with the exception of the `address` property, return `null` unless the corresponding property on all the member ranges has the same value.</span></span>
- <span data-ttu-id="f2b5a-260">A propriedade `address` retorna uma cadeia de caracteres delimitada por vírgulas dos endereços e intervalos dos membros.</span><span class="sxs-lookup"><span data-stu-id="f2b5a-260">The `address` property returns a comma-delimited string of the addresses of the member ranges.</span></span>

<span data-ttu-id="f2b5a-261">Por exemplo, o código a seguir cria um `RangeAreas` no qual apenas um intervalo é uma coluna inteira e apenas um é preenchido com rosa.</span><span class="sxs-lookup"><span data-stu-id="f2b5a-261">For example, the following code creates a `RangeAreas` in which only one range is an entire column and only one is filled with pink.</span></span> <span data-ttu-id="f2b5a-262">O console mostrará `null` para a cor de preenchimento `false` para a propriedade `isEntireRow` e "Planilha1! F3:F5, Planilha1! H:H"(supondo que o nome da planilha  seja "Planilha1") para a propriedade`address`.</span><span class="sxs-lookup"><span data-stu-id="f2b5a-262">The console will show `null` for the fill color, `false` for the `isEntireRow` property, and "Sheet1!F3:F5, Sheet1!H:H" (assuming the sheet name is "Sheet1") for the `address` property.</span></span> 

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

## <a name="see-also"></a><span data-ttu-id="f2b5a-263">Confira também</span><span class="sxs-lookup"><span data-stu-id="f2b5a-263">See also</span></span>

- [<span data-ttu-id="f2b5a-264">Conceitos fundamentais de programação com a API JavaScript do Excel</span><span class="sxs-lookup"><span data-stu-id="f2b5a-264">Fundamental programming concepts with the Excel JavaScript API</span></span>](https://docs.microsoft.com/office/dev/add-ins/reference/overview/excel-add-ins-reference-overview)
- [<span data-ttu-id="f2b5a-265">Objeto Range (API JavaScript para Excel)</span><span class="sxs-lookup"><span data-stu-id="f2b5a-265">Range Object (JavaScript API for Excel)</span></span>](https://docs.microsoft.com/javascript/api/excel/excel.range)
- <span data-ttu-id="f2b5a-266">[Objeto RangeAreas (API JavaScript do Excel)](https://docs.microsoft.com/javascript/api/excel/excel.rangeareas) (esse link pode não funcionar enquanto a API está na visualização.</span><span class="sxs-lookup"><span data-stu-id="f2b5a-266">[RangeAreas Object (JavaScript API for Excel)](https://docs.microsoft.com/javascript/api/excel/excel.rangeareas) (This link may not work while the API is in preview.</span></span> <span data-ttu-id="f2b5a-267">Como alternativa, confira [beta office.d.ts](https://appsforoffice.microsoft.com/lib/beta/hosted/office.d.ts).)</span><span class="sxs-lookup"><span data-stu-id="f2b5a-267">As an alternative, see [beta office.d.ts](https://appsforoffice.microsoft.com/lib/beta/hosted/office.d.ts).)</span></span>