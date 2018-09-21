---
title: Trabalhar com vários intervalos simultaneamente em suplementos do Excel
description: ''
ms.date: 9/4/2018
ms.openlocfilehash: bcb14d1f4c015fe675c2d65cb5f1198d485dd4c5
ms.sourcegitcommit: 3da2038e827dc3f274d63a01dc1f34c98b04557e
ms.translationtype: HT
ms.contentlocale: pt-BR
ms.lasthandoff: 09/19/2018
ms.locfileid: "24016455"
---
# <a name="work-with-multiple-ranges-simultaneously-in-excel-add-ins-preview"></a><span data-ttu-id="a4a8a-102">Trabalhar com vários intervalos simultaneamente em Excel suplementos (Versão Prévia)</span><span class="sxs-lookup"><span data-stu-id="a4a8a-102">Work with multiple ranges simultaneously in Excel add-ins (Preview)</span></span>

<span data-ttu-id="a4a8a-103">A biblioteca JavaScript do Excel permite ao suplemento executar operações e definir propriedades em vários intervalos simultaneamente.</span><span class="sxs-lookup"><span data-stu-id="a4a8a-103">The Excel JavaScript library enables your add-in to perform operations, and set properties, on multiple ranges simultaneously.</span></span> <span data-ttu-id="a4a8a-104">Os intervalos não precisam ser contíguos.</span><span class="sxs-lookup"><span data-stu-id="a4a8a-104">The ranges do not have to be contiguous.</span></span> <span data-ttu-id="a4a8a-105">Além de tornar o seu código mais simples, esta maneira de configurar uma propriedade é executada de forma muito mais rápida do que configurar a mesma propriedade individualmente para cada um dos intervalos.</span><span class="sxs-lookup"><span data-stu-id="a4a8a-105">In addition to making your code simpler, this way of setting a property runs much faster than setting the same property individually for each of the ranges.</span></span>

> [!NOTE]
> <span data-ttu-id="a4a8a-106">As APIs descritas neste artigo exigem a **versão Office 2016 Click-to-Run 1809 Build 10820.20000** ou posterior.</span><span class="sxs-lookup"><span data-stu-id="a4a8a-106">The APIs described in this article require **Office 2016 Click-to-Run version 1809 Build 10820.20000** or later.</span></span> <span data-ttu-id="a4a8a-107">(Talvez você precise ingressar no [programa Office Insider](https://products.office.com/office-insider) para obter uma compilação apropriada.) Além disso, você deve carregar a versão beta da biblioteca do Office JavaScript do [Office. js CDN](https://appsforoffice.microsoft.com/lib/beta/hosted/office.js).</span><span class="sxs-lookup"><span data-stu-id="a4a8a-107">(You may need to join the [Office Insider program](https://products.office.com/office-insider) to get an appropriate build.) Also, you must load the beta version of the Office JavaScript library from [Office.js CDN](https://appsforoffice.microsoft.com/lib/beta/hosted/office.js).</span></span> <span data-ttu-id="a4a8a-108">Por fim, ainda não temos páginas de referência para essas APIs.</span><span class="sxs-lookup"><span data-stu-id="a4a8a-108">Finally, we don't have reference pages for these APIs yet.</span></span> <span data-ttu-id="a4a8a-109">Mas o arquivo de definição a seguir tem descrições para elas: [beta office.d.ts](https://appsforoffice.microsoft.com/lib/beta/hosted/office.d.ts).</span><span class="sxs-lookup"><span data-stu-id="a4a8a-109">But the following definition type file has descriptions for them: [beta office.d.ts](https://appsforoffice.microsoft.com/lib/beta/hosted/office.d.ts).</span></span>

## <a name="rangeareas"></a><span data-ttu-id="a4a8a-110">RangeAreas</span><span class="sxs-lookup"><span data-stu-id="a4a8a-110">RangeAreas</span></span>

<span data-ttu-id="a4a8a-111">Um conjunto de intervalos (possivelmente não adjacentes) é representado por um objeto `Excel.RangeAreas` .</span><span class="sxs-lookup"><span data-stu-id="a4a8a-111">A set of (possibly discontiguous) ranges is represented by an `Excel.RangeAreas` object.</span></span> <span data-ttu-id="a4a8a-112">Ele tem propriedades e métodos semelhantes ao tipo `Range` (vários com nomes semelhantes ou iguais), contudo, ajustes foram feitos nos:</span><span class="sxs-lookup"><span data-stu-id="a4a8a-112">It has properties and methods similar to the `Range` type (many with the same, or similar, names), but adjustments have been made to:</span></span>

- <span data-ttu-id="a4a8a-113">Tipos de dados para as propriedades e no comportamento dos setters e getters.</span><span class="sxs-lookup"><span data-stu-id="a4a8a-113">The data types for properties and the behavior of the setters and getters.</span></span>
- <span data-ttu-id="a4a8a-114">Tipos de dados dos parâmetros do método e nos comportamentos do método.</span><span class="sxs-lookup"><span data-stu-id="a4a8a-114">The data types of method parameters and the method behaviors.</span></span>
- <span data-ttu-id="a4a8a-115">Valores retornados dos tipos de dados do método.</span><span class="sxs-lookup"><span data-stu-id="a4a8a-115">The data types of method return values.</span></span>

<span data-ttu-id="a4a8a-116">Alguns exemplos:</span><span class="sxs-lookup"><span data-stu-id="a4a8a-116">Some examples:</span></span>

- <span data-ttu-id="a4a8a-117">`RangeAreas` tem uma propriedade  `address` que retorna uma sequência de caracteres delimitada por vírgula do intervalo de endereços, em vez de apenas um endereço como na propriedade `Range.address` .</span><span class="sxs-lookup"><span data-stu-id="a4a8a-117">`RangeAreas` has an `address` property that returns a comma-delimited string of range addresses, instead of just one address as with the `Range.address` property.</span></span>
- <span data-ttu-id="a4a8a-118">`RangeAreas` tem uma propriedade `dataValidation` que retorna um objeto  `DataValidation` que representa a validação de dados de todos os intervalos no `RangeAreas`, caso seja consistente.</span><span class="sxs-lookup"><span data-stu-id="a4a8a-118">`RangeAreas` has a `dataValidation` property that returns a `DataValidation` object that represents the data validation of all the ranges in the `RangeAreas`, if it is consistent.</span></span> <span data-ttu-id="a4a8a-119">A propriedade é `null` se objetos `DataValidation` idênticos não forem aplicados a todos os intervalos em `RangeAreas`.</span><span class="sxs-lookup"><span data-stu-id="a4a8a-119">The property is `null` if identical `DataValidation` objects are not applied to all the all the ranges in the `RangeAreas`.</span></span> <span data-ttu-id="a4a8a-120">Esse é um princípio geral, mas não universal, com o objeto `RangeAreas` : *se uma propriedade não tiver valores consistentes em todos os intervalos no `RangeAreas`, então ela é `null`.*</span><span class="sxs-lookup"><span data-stu-id="a4a8a-120">This is a general, but not universal, principle with the `RangeAreas` object: *If a property does not have consistent values on all the all the ranges in the `RangeAreas`, then it is `null`.*</span></span> <span data-ttu-id="a4a8a-121">Confira [Propriedades de leitura de RangeAreas](#reading-properties-of-rangeareas) para obter mais informações e conhecer algumas exceções.</span><span class="sxs-lookup"><span data-stu-id="a4a8a-121">See [Reading properties of RangeAreas](#reading-properties-of-rangeareas) for more information and some exceptions.</span></span>
- <span data-ttu-id="a4a8a-122">`RangeAreas.cellCount` obtém o número total de células em todos os intervalos de `RangeAreas`.</span><span class="sxs-lookup"><span data-stu-id="a4a8a-122">`RangeAreas.cellCount` gets the total number of cells in all the ranges in the `RangeAreas`.</span></span>
- <span data-ttu-id="a4a8a-123">`RangeAreas.calculate` recalcula as células de todos os intervalos de `RangeAreas`.</span><span class="sxs-lookup"><span data-stu-id="a4a8a-123">`RangeAreas.calculate` recalculates the cells of all the ranges in the `RangeAreas`.</span></span>
- <span data-ttu-id="a4a8a-124">`RangeAreas.getEntireColumn` e `RangeAreas.getEntireRow` retorna outro objeto `RangeAreas` que representa todas as colunas (ou linhas) em todos os intervalos de `RangeAreas`.</span><span class="sxs-lookup"><span data-stu-id="a4a8a-124">`RangeAreas.getEntireColumn` and `RangeAreas.getEntireRow` return another `RangeAreas` object that represents all of the columns (or rows) in all the ranges in the `RangeAreas`.</span></span> <span data-ttu-id="a4a8a-125">Por exemplo, se o `RangeAreas` representa "A1: C4" e "F14:L15", então `RangeAreas.getEntireColumn` retorna um objeto `RangeAreas` que representa "A:C" e "F:L".</span><span class="sxs-lookup"><span data-stu-id="a4a8a-125">For example, if the `RangeAreas` represents "A1:C4" and "F14:L15", then `RangeAreas.getEntireColumn` returns a `RangeAreas` object that represents "A:C" and "F:L".</span></span>
- <span data-ttu-id="a4a8a-126">`RangeAreas.copyFrom` pode receber um parâmetro `Range` ou `RangeAreas` que representa o(s) intervalo(s) de origem da operação de cópia.</span><span class="sxs-lookup"><span data-stu-id="a4a8a-126">`RangeAreas.copyFrom` can take either a `Range` or a `RangeAreas` parameter representing the source range(s) of the copy operation.</span></span>

#### <a name="complete-list-of-range-members-that-are-also-available-on-rangeareas"></a><span data-ttu-id="a4a8a-127">Lista completa dos membros de Range que também estão disponíveis em RangeAreas</span><span class="sxs-lookup"><span data-stu-id="a4a8a-127">Complete list of Range members that are also available on RangeAreas</span></span>

##### <a name="properties"></a><span data-ttu-id="a4a8a-128">Propriedades</span><span class="sxs-lookup"><span data-stu-id="a4a8a-128">Properties</span></span>

<span data-ttu-id="a4a8a-129">Esteja familiarizado com [as propriedades de leitura de RangeAreas](#reading-properties-of-rangeareas) antes de escrever código para ler as propriedades listadas.</span><span class="sxs-lookup"><span data-stu-id="a4a8a-129">Be familiar with [Reading properties of RangeAreas](#reading-properties-of-rangeareas) before you write code that reads any properties listed.</span></span> <span data-ttu-id="a4a8a-130">Há sutilezas para o que é retornado.</span><span class="sxs-lookup"><span data-stu-id="a4a8a-130">There are subtleties to what gets returned.</span></span>

- <span data-ttu-id="a4a8a-131">address</span><span class="sxs-lookup"><span data-stu-id="a4a8a-131">address</span></span>
- <span data-ttu-id="a4a8a-132">addressLocal</span><span class="sxs-lookup"><span data-stu-id="a4a8a-132">addressLocal</span></span>
- <span data-ttu-id="a4a8a-133">cellCount</span><span class="sxs-lookup"><span data-stu-id="a4a8a-133">cellCount</span></span>
- <span data-ttu-id="a4a8a-134">conditionalFormats</span><span class="sxs-lookup"><span data-stu-id="a4a8a-134">conditionalFormats</span></span>
- <span data-ttu-id="a4a8a-135">context</span><span class="sxs-lookup"><span data-stu-id="a4a8a-135">context</span></span>
- <span data-ttu-id="a4a8a-136">dataValidation</span><span class="sxs-lookup"><span data-stu-id="a4a8a-136">dataValidation</span></span>
- <span data-ttu-id="a4a8a-137">format</span><span class="sxs-lookup"><span data-stu-id="a4a8a-137">format</span></span>
- <span data-ttu-id="a4a8a-138">isEntireColumn</span><span class="sxs-lookup"><span data-stu-id="a4a8a-138">isEntireColumn</span></span>
- <span data-ttu-id="a4a8a-139">isEntireRow</span><span class="sxs-lookup"><span data-stu-id="a4a8a-139">isEntireRow</span></span>
- <span data-ttu-id="a4a8a-140">style</span><span class="sxs-lookup"><span data-stu-id="a4a8a-140">style</span></span>
- <span data-ttu-id="a4a8a-141">worksheet</span><span class="sxs-lookup"><span data-stu-id="a4a8a-141">worksheet</span></span>

##### <a name="methods"></a><span data-ttu-id="a4a8a-142">Métodos</span><span class="sxs-lookup"><span data-stu-id="a4a8a-142">Methods</span></span>

<span data-ttu-id="a4a8a-143">Métodos de Range em versão prévia estão marcados.</span><span class="sxs-lookup"><span data-stu-id="a4a8a-143">Range methods in preview are marked.</span></span>

- <span data-ttu-id="a4a8a-144">calculate()</span><span class="sxs-lookup"><span data-stu-id="a4a8a-144">calculate()</span></span>
- <span data-ttu-id="a4a8a-145">clear()</span><span class="sxs-lookup"><span data-stu-id="a4a8a-145">clear()</span></span>
- <span data-ttu-id="a4a8a-146">convertDataTypeToText() (versão prévia)</span><span class="sxs-lookup"><span data-stu-id="a4a8a-146">convertDataTypeToText() (preview)</span></span>
- <span data-ttu-id="a4a8a-147">convertToLinkedDataType() (versão prévia)</span><span class="sxs-lookup"><span data-stu-id="a4a8a-147">convertToLinkedDataType() (preview)</span></span>
- <span data-ttu-id="a4a8a-148">copyFrom() (versão prévia)</span><span class="sxs-lookup"><span data-stu-id="a4a8a-148">copyFrom() (preview)</span></span>
- <span data-ttu-id="a4a8a-149">getEntireColumn()</span><span class="sxs-lookup"><span data-stu-id="a4a8a-149">getEntireColumn()</span></span>
- <span data-ttu-id="a4a8a-150">getEntireRow()</span><span class="sxs-lookup"><span data-stu-id="a4a8a-150">getEntireRow()</span></span>
- <span data-ttu-id="a4a8a-151">getIntersection()</span><span class="sxs-lookup"><span data-stu-id="a4a8a-151">getIntersection()</span></span>
- <span data-ttu-id="a4a8a-152">getIntersectionOrNullObject()</span><span class="sxs-lookup"><span data-stu-id="a4a8a-152">getIntersectionOrNullObject()</span></span>
- <span data-ttu-id="a4a8a-153">getOffsetRange() (chamado getOffsetRangeAreas no objeto RangeAreas)</span><span class="sxs-lookup"><span data-stu-id="a4a8a-153">getOffsetRange() (named getOffsetRangeAreas on the RangeAreas object)</span></span>
- <span data-ttu-id="a4a8a-154">getSpecialCells() (versão prévia)</span><span class="sxs-lookup"><span data-stu-id="a4a8a-154">getSpecialCells() (preview)</span></span>
- <span data-ttu-id="a4a8a-155">getSpecialCellsOrNullObject() (versão prévia)</span><span class="sxs-lookup"><span data-stu-id="a4a8a-155">getSpecialCellsOrNullObject() (preview)</span></span>
- <span data-ttu-id="a4a8a-156">getTables() (versão prévia)</span><span class="sxs-lookup"><span data-stu-id="a4a8a-156">getTables() (preview)</span></span>
- <span data-ttu-id="a4a8a-157">getUsedRange() (chamado getUsedRangeAreas no objeto RangeAreas)</span><span class="sxs-lookup"><span data-stu-id="a4a8a-157">getUsedRange() (named getUsedRangeAreas on the RangeAreas object)</span></span>
- <span data-ttu-id="a4a8a-158">getUsedRangeOrNullObject() (chamado getUsedRangeAreasOrNullObject no objeto RangeAreas)</span><span class="sxs-lookup"><span data-stu-id="a4a8a-158">getUsedRangeOrNullObject() (named getUsedRangeAreasOrNullObject on the RangeAreas object)</span></span>
- <span data-ttu-id="a4a8a-159">load()</span><span class="sxs-lookup"><span data-stu-id="a4a8a-159">load()</span></span>
- <span data-ttu-id="a4a8a-160">set()</span><span class="sxs-lookup"><span data-stu-id="a4a8a-160">set\*</span></span>
- <span data-ttu-id="a4a8a-161">setDirty() (versão prévia)</span><span class="sxs-lookup"><span data-stu-id="a4a8a-161">setDirty() (preview)</span></span>
- <span data-ttu-id="a4a8a-162">toJSON()</span><span class="sxs-lookup"><span data-stu-id="a4a8a-162">toJSON()</span></span>
- <span data-ttu-id="a4a8a-163">track()</span><span class="sxs-lookup"><span data-stu-id="a4a8a-163">track</span></span>
- <span data-ttu-id="a4a8a-164">untrack()</span><span class="sxs-lookup"><span data-stu-id="a4a8a-164">untrack()</span></span>

### <a name="rangearea-specific-properties-and-methods"></a><span data-ttu-id="a4a8a-165">Propriedades e métodos específicos de RangeArea</span><span class="sxs-lookup"><span data-stu-id="a4a8a-165">RangeArea-specific properties and methods</span></span>

<span data-ttu-id="a4a8a-166">O tipo `RangeAreas` tem algumas propriedades e métodos que não estão no objeto`Range`.</span><span class="sxs-lookup"><span data-stu-id="a4a8a-166">The `RangeAreas` type has some properties and methods that are not on the `Range` object.</span></span> <span data-ttu-id="a4a8a-167">A seguir está uma seleção deles:</span><span class="sxs-lookup"><span data-stu-id="a4a8a-167">The following is a selection of them:</span></span>

- <span data-ttu-id="a4a8a-168">`areas`: Um objeto `RangeCollection` que contém todos os intervalos representados pelo objeto `RangeAreas`.</span><span class="sxs-lookup"><span data-stu-id="a4a8a-168">`areas`: A `RangeCollection` object that contains all of the ranges represented by the `RangeAreas` object.</span></span> <span data-ttu-id="a4a8a-169">O objeto  `RangeCollection` também é novo e é similar a outros objetos da coleção do Excel.</span><span class="sxs-lookup"><span data-stu-id="a4a8a-169">The `RangeCollection` object is also new and is similar to other Excel collection objects.</span></span> <span data-ttu-id="a4a8a-170">Ele tem uma propriedade `items` que é uma matriz de objetos `Range` que representa os intervalos.</span><span class="sxs-lookup"><span data-stu-id="a4a8a-170">It has an `items` property which is an array of `Range` objects representing the ranges.</span></span>
- <span data-ttu-id="a4a8a-171">`areaCount`: O número total de intervalos no `RangeAreas`.</span><span class="sxs-lookup"><span data-stu-id="a4a8a-171">`areaCount`: The total number of ranges in the `RangeAreas`.</span></span>
- <span data-ttu-id="a4a8a-172">`getOffsetRangeAreas`: Funciona exatamente como [Range.getOffsetRange](https://docs.microsoft.com/javascript/api/excel/excel.range#getoffsetrange-rowoffset--columnoffset-), exceto que um `RangeAreas` é retornado e contém intervalos que são um deslocamento de um dos intervalos no `RangeAreas` original.</span><span class="sxs-lookup"><span data-stu-id="a4a8a-172">`getOffsetRangeAreas`: Works just like [Range.getOffsetRange](https://docs.microsoft.com/javascript/api/excel/excel.range#getoffsetrange-rowoffset--columnoffset-), except that a `RangeAreas` is returned and it contains ranges that are each offset from one of the ranges in the original `RangeAreas`.</span></span>

## <a name="create-rangeareas-and-set-properties"></a><span data-ttu-id="a4a8a-173">Criar RangeAreas e definir propriedades</span><span class="sxs-lookup"><span data-stu-id="a4a8a-173">Create RangeAreas and set properties</span></span>

<span data-ttu-id="a4a8a-174">Você pode criar o objeto  `RangeAreas` de duas formas básicas:</span><span class="sxs-lookup"><span data-stu-id="a4a8a-174">You can create `RangeAreas` object in two basic ways:</span></span>

- <span data-ttu-id="a4a8a-175">Chamar `Worksheet.getRanges()` e passar a ele uma sequência de caracteres com um intervalo de endereços delimitados por vírgula.</span><span class="sxs-lookup"><span data-stu-id="a4a8a-175">Call `Worksheet.getRanges()` and pass it a string with comma-delimited range addresses.</span></span> <span data-ttu-id="a4a8a-176">Se algum intervalo que você deseja incluir tiver sido transformado em [NamedItem](https://docs.microsoft.com/javascript/api/excel/excel.nameditem), você pode incluir o nome, em vez do endereço, na sequência de caracteres.</span><span class="sxs-lookup"><span data-stu-id="a4a8a-176">If any range you want to include has been made into a [NamedItem](https://docs.microsoft.com/javascript/api/excel/excel.nameditem), you can include the name, instead of the address, in the string.</span></span>
- <span data-ttu-id="a4a8a-177">Chame `Workbook.getSelectedRanges()`.</span><span class="sxs-lookup"><span data-stu-id="a4a8a-177">Call `Workbook.getSelectedRanges()`.</span></span> <span data-ttu-id="a4a8a-178">Esse método retorna um `RangeAreas` que representa todos os intervalos selecionados na planilha ativa no momento.</span><span class="sxs-lookup"><span data-stu-id="a4a8a-178">This method returns a `RangeAreas` representing all the ranges that are selected on the currently active worksheet.</span></span>

<span data-ttu-id="a4a8a-179">Depois que você tiver um objeto `RangeAreas` , você pode criar outros usando os métodos no objeto que retornam `RangeAreas` como `getOffsetRangeAreas` e `getIntersection`.</span><span class="sxs-lookup"><span data-stu-id="a4a8a-179">Once you have a `RangeAreas` object, you can create others using the methods on the object that return `RangeAreas` such as `getOffsetRangeAreas` and `getIntersection`.</span></span>

> [!NOTE]
> <span data-ttu-id="a4a8a-180">Você não pode adicionar diretamente intervalos adicionais para um objeto `RangeAreas` .</span><span class="sxs-lookup"><span data-stu-id="a4a8a-180">You cannot directly add additional ranges to a `RangeAreas` object.</span></span> <span data-ttu-id="a4a8a-181">Por exemplo, a coleção em `RangeAreas.areas` não tem um método `add` .</span><span class="sxs-lookup"><span data-stu-id="a4a8a-181">For example, the collection in `RangeAreas.areas` does not have an `add` method.</span></span>


> [!WARNING] 
> <span data-ttu-id="a4a8a-182">Não tente adicionar ou excluir membros diretamente na matriz `RangeAreas.areas.items` .</span><span class="sxs-lookup"><span data-stu-id="a4a8a-182">Do not attempt to directly add or delete members of the the `RangeAreas.areas.items` array.</span></span> <span data-ttu-id="a4a8a-183">Isso levará a um comportamento indesejável em seu código.</span><span class="sxs-lookup"><span data-stu-id="a4a8a-183">This will lead to undesirable behavior in your code.</span></span> <span data-ttu-id="a4a8a-184">Por exemplo, é possível adicionar um objeto  `Range` na matriz, mas isso causará erros, porque os métodos e propriedades `RangeAreas` se comportarão como se o novo item não estivesse lá.</span><span class="sxs-lookup"><span data-stu-id="a4a8a-184">For example, it is possible to push an additional `Range` object onto the array, but doing so will cause errors because `RangeAreas` properties and methods behave as if the new item isn't there.</span></span> <span data-ttu-id="a4a8a-185">Por exemplo, a propriedade  `areaCount` não inclui intervalos adicionados dessa maneira e o `RangeAreas.getItemAt(index)` gera um erro se `index` for maior do que `areasCount-1`.</span><span class="sxs-lookup"><span data-stu-id="a4a8a-185">For example, the `areaCount` property does not include ranges pushed in this way, and the `RangeAreas.getItemAt(index)` throws an error if `index` is larger than `areasCount-1`.</span></span> <span data-ttu-id="a4a8a-186">Da mesma forma, excluir um objeto `Range` na matriz `RangeAreas.areas.items` obtendo uma referência a ele e chamando seu método `Range.delete` causará bugs: embora o objeto `Range` *esteja* excluído, as propriedades e os métodos do objeto pai `RangeAreas` se comportam, ou tentam, como se ele ainda existisse.</span><span class="sxs-lookup"><span data-stu-id="a4a8a-186">Similarly, deleting a `Range` object in the `RangeAreas.areas.items` array by getting a reference to it and calling its `Range.delete` method causes bugs: although the `Range` object *is* deleted, the properties and methods of the parent `RangeAreas` object behave, or try to, as if it is still in existence.</span></span> <span data-ttu-id="a4a8a-187">Por exemplo, se o seu código chama `RangeAreas.calculate`, o Office tentará calcular o intervalo, mas apresentará um erro, porque o objeto do intervalo não existirá.</span><span class="sxs-lookup"><span data-stu-id="a4a8a-187">For example, if your code calls `RangeAreas.calculate`, Office will try to calculate the range, but will error because the range object is gone.</span></span>

<span data-ttu-id="a4a8a-188">Configurar a propriedade em um `RangeAreas` define a propriedade correspondente em todos os intervalos na coleção `RangeAreas.areas` .</span><span class="sxs-lookup"><span data-stu-id="a4a8a-188">Setting a property on a `RangeAreas` sets the corresponding property on all the ranges in the `RangeAreas.areas` collection.</span></span>

<span data-ttu-id="a4a8a-189">A seguir está um exemplo sobre configuração de propriedade em vários intervalos.</span><span class="sxs-lookup"><span data-stu-id="a4a8a-189">The following is an example of setting a property on multiple ranges.</span></span> <span data-ttu-id="a4a8a-190">A função realça os intervalos **F3:F5** e **H3:H5**.</span><span class="sxs-lookup"><span data-stu-id="a4a8a-190">The function highlights the ranges **F3:F5** and **H3:H5**.</span></span>

```js
Excel.run(function (context) {
    const sheet = context.workbook.worksheets.getActiveWorksheet();
    const rangeAreas = sheet.getRanges("F3:F5, H3:H5");
    rangeAreas.format.fill.color = "pink";

    return context.sync();
})
```

<span data-ttu-id="a4a8a-191">Este exemplo se aplica a cenários nos quais você pode codificar os endereços de intervalo que você passa para `getRanges` ou facilmente os calcula em tempo de execução.</span><span class="sxs-lookup"><span data-stu-id="a4a8a-191">This example applies to scenarios in which you can hard code the range addresses that you pass to `getRanges` or easily calculate them at runtime.</span></span> <span data-ttu-id="a4a8a-192">Alguns dos cenários em que isso seria verdadeiro são:</span><span class="sxs-lookup"><span data-stu-id="a4a8a-192">Some of the scenarios in which this would be true include:</span></span> 

- <span data-ttu-id="a4a8a-193">O código é executado no contexto de um modelo conhecido.</span><span class="sxs-lookup"><span data-stu-id="a4a8a-193">The code runs in the context of a known template.</span></span>
- <span data-ttu-id="a4a8a-194">O código é executado no contexto de dados importados onde o esquema dos dados é conhecido.</span><span class="sxs-lookup"><span data-stu-id="a4a8a-194">The code runs in the context of imported data where the schema of the data is known.</span></span>

<span data-ttu-id="a4a8a-195">Quando você não sabe em tempo de codificação quais intervalos que você precisa operar, você deve descobri-los em tempo de execução.</span><span class="sxs-lookup"><span data-stu-id="a4a8a-195">When you can't know at coding-time which ranges you need to operate on, you must discover them at runtime.</span></span> <span data-ttu-id="a4a8a-196">A próxima seção discute esses cenários.</span><span class="sxs-lookup"><span data-stu-id="a4a8a-196">The next section discusses these scenarios.</span></span>

### <a name="discover-range-areas-programmatically"></a><span data-ttu-id="a4a8a-197">Descobrir áreas de intervalo programaticamente</span><span class="sxs-lookup"><span data-stu-id="a4a8a-197">Discover range areas programmatically</span></span>

<span data-ttu-id="a4a8a-198">Os métodos `Range.getSpecialCells()` e `Range.getSpecialCellsOrNullObject()` permitem que você encontre em tempo de execução os intervalos que você deseja operar com base nas características das células e no tipo dos valores das células.</span><span class="sxs-lookup"><span data-stu-id="a4a8a-198">The `Range.getSpecialCells()` and `Range.getSpecialCellsOrNullObject()` methods enable you to find at runtime the ranges that you want to operate on based on the characteristics of the cells and the type of the values of the cells.</span></span> <span data-ttu-id="a4a8a-199">Aqui estão as assinaturas dos métodos dos arquivos de dados TypeScript:</span><span class="sxs-lookup"><span data-stu-id="a4a8a-199">Here are the signatures of the methods from the TypeScript data types file:</span></span>

```typescript
getSpecialCells(cellType: Excel.SpecialCellType, cellValueType?: Excel.SpecialCellValueType): Excel.RangeAreas;
```

```typescript
getSpecialCellsOrNullObject(cellType: Excel.SpecialCellType, cellValueType?: Excel.SpecialCellValueType): Excel.RangeAreas;
```

<span data-ttu-id="a4a8a-200">A seguir está um exemplo do primeiro caso.</span><span class="sxs-lookup"><span data-stu-id="a4a8a-200">The following is an example of using the "Between" operator:</span></span> <span data-ttu-id="a4a8a-201">Sobre este código, observe:</span><span class="sxs-lookup"><span data-stu-id="a4a8a-201">About this code, note:</span></span>

- <span data-ttu-id="a4a8a-202">Ele limita a parte da planilha que precisa ser pesquisada chamando primeiro `Worksheet.getUsedRange` e chamando `getSpecialCells` somente para aquele intervalo.</span><span class="sxs-lookup"><span data-stu-id="a4a8a-202">It limits the part of the sheet that needs to be searched by first calling `Worksheet.getUsedRange` and calling `getSpecialCells` for only that range.</span></span>
- <span data-ttu-id="a4a8a-203">Ele passa em forma de parâmetro para `getSpecialCells`, a versão da sequência de caracteres de um valor da enumeração `Excel.SpecialCellType` .</span><span class="sxs-lookup"><span data-stu-id="a4a8a-203">It passes as a parameter to `getSpecialCells` the string version of a value from the `Excel.SpecialCellType` enum.</span></span> <span data-ttu-id="a4a8a-204">Alguns outros valores que também podem ser passados são "Blanks" para células vazias, "Constants" para células com valores literais em vez de fórmulas e "SameConditionalFormat" para células que possuem a mesma formatação condicional como a primeira célula no `usedRange`.</span><span class="sxs-lookup"><span data-stu-id="a4a8a-204">Some of the other values that could be passed instead are "Blanks" for empty cells, "Constants" for cells with literal values instead of formulas, and "SameConditionalFormat" for cells that have the same conditional formatting as the first cell in the `usedRange`.</span></span> <span data-ttu-id="a4a8a-205">A primeira célula é a primeira célula no canto esquerdo superior.</span><span class="sxs-lookup"><span data-stu-id="a4a8a-205">The first cell is the upper leftmost cell.</span></span> <span data-ttu-id="a4a8a-206">Para obter uma lista completa dos valores na enumeração, consulte [beta office.d.ts](https://appsforoffice.microsoft.com/lib/beta/hosted/office.d.ts).</span><span class="sxs-lookup"><span data-stu-id="a4a8a-206">For a complete list of the values in the enum, see [beta office.d.ts](https://appsforoffice.microsoft.com/lib/beta/hosted/office.d.ts).</span></span>
- <span data-ttu-id="a4a8a-207">O método `getSpecialCells` retorna um objeto `RangeAreas`, portanto todas as células com fórmulas serão cor de rosa, mesmo que não sejam contíguas.</span><span class="sxs-lookup"><span data-stu-id="a4a8a-207">The `getSpecialCells` method returns a `RangeAreas` object, so all of the cells with formulas will be colored pink even if they are not all contiguous.</span></span> 

```js
Excel.run(function (context) {
    const sheet = context.workbook.worksheets.getActiveWorksheet();
    const usedRange = sheet.getUsedRange();
    const formulaRanges = usedRange.getSpecialCells("Formulas");
    formulaRanges.format.fill.color = "pink";

    return context.sync();
})
```

<span data-ttu-id="a4a8a-208">Às vezes, o intervalo não tem *nenhuma* célula com a característica procurada.</span><span class="sxs-lookup"><span data-stu-id="a4a8a-208">Sometimes the range doesn't have *any* cells with the targeted characteristic.</span></span> <span data-ttu-id="a4a8a-209">Se `getSpecialCells` não encontrar nada, ele gera um erro de **ItemNotFound**.</span><span class="sxs-lookup"><span data-stu-id="a4a8a-209">If `getSpecialCells` doesn't find any, it throws an **ItemNotFound** error.</span></span> <span data-ttu-id="a4a8a-210">Isso desviaria o fluxo de controle para um bloco/método `catch` , se houver um.</span><span class="sxs-lookup"><span data-stu-id="a4a8a-210">This would divert the flow of control to a `catch` block/method, if there is one.</span></span> <span data-ttu-id="a4a8a-211">Se não houver, o erro interrompe a função.</span><span class="sxs-lookup"><span data-stu-id="a4a8a-211">If there isn't, the error halts the function.</span></span> <span data-ttu-id="a4a8a-212">Pode haver cenários nos quais gerar um erro é exatamente o que você deseja quando não há nenhuma célula com a característica alvo.</span><span class="sxs-lookup"><span data-stu-id="a4a8a-212">There may be scenarios in which throwing the error is exactly what you want to happen when there are no cells with the targeted characteristic.</span></span> 

<span data-ttu-id="a4a8a-213">Mas em cenários onde é normal, mas talvez incomum, não existir nenhuma célula correspondente, seu código deve verificar essa possibilidade e gerencia-la normalmente sem exibir um erro.</span><span class="sxs-lookup"><span data-stu-id="a4a8a-213">But in scenarios in which it is normal, but perhaps uncommon, for there to be no matching cells; your code should check for this possibility and handle it gracefully without throwing an error.</span></span> <span data-ttu-id="a4a8a-214">Para esses cenários, use o método `getSpecialCellsOrNullObject` e teste a propriedade `RangeAreas.isNullObject`.</span><span class="sxs-lookup"><span data-stu-id="a4a8a-214">For these scenarios, use the `getSpecialCellsOrNullObject` method and test the `RangeAreas.isNullObject` property.</span></span> <span data-ttu-id="a4a8a-215">Apresentamos um exemplo a seguir.</span><span class="sxs-lookup"><span data-stu-id="a4a8a-215">The following is an example.</span></span> <span data-ttu-id="a4a8a-216">Observação sobre o código:</span><span class="sxs-lookup"><span data-stu-id="a4a8a-216">Note about this code:</span></span>

- <span data-ttu-id="a4a8a-217">O método `getSpecialCellsOrNullObject` sempre retorna um objeto de proxy, portanto, ele nunca é `null` no sentido comum do JavaScript.</span><span class="sxs-lookup"><span data-stu-id="a4a8a-217">The `getSpecialCellsOrNullObject` method always returns a proxy object, so it is never `null` in the ordinary JavaScript sense.</span></span> <span data-ttu-id="a4a8a-218">Mas se nenhuma célula correspondente for encontrada, a propriedade  `isNullObject` do objeto é definida como `true`.</span><span class="sxs-lookup"><span data-stu-id="a4a8a-218">But if no matching cells are found, the `isNullObject` property of the object is set to `true`.</span></span>
- <span data-ttu-id="a4a8a-219">Ele chama `context.sync` *antes* de testar a propriedade `isNullObject` .</span><span class="sxs-lookup"><span data-stu-id="a4a8a-219">It calls `context.sync` *before* it tests the `isNullObject` property.</span></span> <span data-ttu-id="a4a8a-220">Esse é um requisito com todas as propriedades e métodos `*OrNullObject`, porque você sempre precisa carregar e sincronizar uma propriedade a fim de lê-la.</span><span class="sxs-lookup"><span data-stu-id="a4a8a-220">This is a requirement with all `*OrNullObject` methods and properties, because you always have to load and sync a property in order to read it.</span></span> <span data-ttu-id="a4a8a-221">No entanto, não é necessário *explicitamente* carregar a propriedade `isNullObject` .</span><span class="sxs-lookup"><span data-stu-id="a4a8a-221">However, it is not necessary to *explicitly* load the `isNullObject` property.</span></span> <span data-ttu-id="a4a8a-222">Ela é carregada automaticamente pelo `context.sync`, mesmo se `load`  não é chamado no objeto.</span><span class="sxs-lookup"><span data-stu-id="a4a8a-222">It is automatically loaded by the `context.sync` even if `load` is not called on the object.</span></span> <span data-ttu-id="a4a8a-223">Para obter mais informações, consulte [\*OrNullObject](https://docs.microsoft.com/office/dev/add-ins/excel/excel-add-ins-advanced-concepts#42ornullobject-methods).</span><span class="sxs-lookup"><span data-stu-id="a4a8a-223">For more information, see [\*](https://docs.microsoft.com/office/dev/add-ins/excel/excel-add-ins-advanced-concepts#42ornullobject-methods).</span></span>
- <span data-ttu-id="a4a8a-224">Você pode testar esse código selecionando um intervalo que tenha células sem fórmulas e o executando.</span><span class="sxs-lookup"><span data-stu-id="a4a8a-224">You can test this code by first selecting a range that has no formula cells and running it.</span></span> <span data-ttu-id="a4a8a-225">Em seguida, selecione um intervalo que tenha pelo menos uma célula com fórmula e o execute novamente.</span><span class="sxs-lookup"><span data-stu-id="a4a8a-225">Then select a range that has at least one cell with a formula and run it again.</span></span>

```js
Excel.run(function (context) {
    const range = context.workbook.getSelectedRange();
    const formulaRanges = range.getSpecialCellsOrNullObject("Formulas");
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

<span data-ttu-id="a4a8a-226">Para manter a simplicidade, todos os outros exemplos neste artigo usam o método `getSpecialCells` em vez de `getSpecialCellsOrNullObject`.</span><span class="sxs-lookup"><span data-stu-id="a4a8a-226">For simplicity, all other examples in this article use the `getSpecialCells` method instead of  `getSpecialCellsOrNullObject`.</span></span>

#### <a name="narrow-the-target-cells-with-cell-value-types"></a><span data-ttu-id="a4a8a-227">Restringir as células de destino com tipos de valores de célula</span><span class="sxs-lookup"><span data-stu-id="a4a8a-227">Narrow the target cells with cell value types</span></span>

<span data-ttu-id="a4a8a-228">Há um segundo parâmetro opcional, do tipo enumerado `Excel.SpecialCellValueType`, que restringe ainda mais as células de destino.</span><span class="sxs-lookup"><span data-stu-id="a4a8a-228">There is an optional second parameter, of enum type `Excel.SpecialCellValueType`, that further narrows down the cells to target.</span></span> <span data-ttu-id="a4a8a-229">Você pode usá-lo apenas quando você passá "Formulas" ou "Constants" para `getSpecialCells` ou `getSpecialCellsOrNullObject`.</span><span class="sxs-lookup"><span data-stu-id="a4a8a-229">You can use it only when you pass either "Formulas" or "Constants" to `getSpecialCells` or `getSpecialCellsOrNullObject`.</span></span> <span data-ttu-id="a4a8a-230">O parâmetro especifica que você apenas deseja as células com determinados tipos de valores.</span><span class="sxs-lookup"><span data-stu-id="a4a8a-230">The parameter specifies that you only want cells with certain types of values.</span></span> <span data-ttu-id="a4a8a-231">Há quatro tipos básicos: "Erro", "Lógico" (que é booleano), "Números" e "Texto".</span><span class="sxs-lookup"><span data-stu-id="a4a8a-231">There are four basic types: "Error", "Logical" (which means boolean), "Numbers", and "Text".</span></span> <span data-ttu-id="a4a8a-232">(A enumeração tem outros valores além desses quatro abordados abaixo). A seguir está um exemplo.</span><span class="sxs-lookup"><span data-stu-id="a4a8a-232">(The enum has other values besides these four which are discussed below.) The following is an example.</span></span> <span data-ttu-id="a4a8a-233">Sobre este código, observe:</span><span class="sxs-lookup"><span data-stu-id="a4a8a-233">About this code, note:</span></span>

- <span data-ttu-id="a4a8a-234">Ele realçará somente células que têm um valor numérico literal.</span><span class="sxs-lookup"><span data-stu-id="a4a8a-234">It will only highlight cells that have a literal number value.</span></span> <span data-ttu-id="a4a8a-235">Ele não irá realçar células que têm uma fórmula (mesmo se o resultado é um número) ou uma célula que contenha um valor booleano, texto ou erro.</span><span class="sxs-lookup"><span data-stu-id="a4a8a-235">It will not highlight cells that have a formula (even if the result is a number) or a boolean, text, or error state cells.</span></span>
- <span data-ttu-id="a4a8a-236">Para testar o código, certifique-se de que a planilha possua algumas células com valores numéricos literais, algumas com outros tipos de valores literais e algumas células com fórmulas.</span><span class="sxs-lookup"><span data-stu-id="a4a8a-236">To test the code, be sure the worksheet has some cells with literal number values, some with other kinds of literal values, and some with formulas.</span></span>

```js
Excel.run(function (context) {
    const sheet = context.workbook.worksheets.getActiveWorksheet();
    const usedRange = sheet.getUsedRange();
    const constantNumberRanges = usedRange.getSpecialCells("Constants", "Numbers");
    constantNumberRanges.format.fill.color = "pink";

    return context.sync();
})
```

<span data-ttu-id="a4a8a-237">Em alguns casos, você precisa operar em mais de um tipo de valor de célula, como valores de texto e valores booleanos ("Lógica").</span><span class="sxs-lookup"><span data-stu-id="a4a8a-237">Sometimes you need to operate on more than one cell value type, such as all text-valued and all boolean-valued ("Logical") cells.</span></span> <span data-ttu-id="a4a8a-238">A enumeração `Excel.SpecialCellValueType` tem valores que permitem que você combine tipos.</span><span class="sxs-lookup"><span data-stu-id="a4a8a-238">The `Excel.SpecialCellValueType` enum has values that let you combine types.</span></span> <span data-ttu-id="a4a8a-239">Por exemplo, "LogicalText" irá marcar todas as células com valores booleanos e todas as células com valores de texto.</span><span class="sxs-lookup"><span data-stu-id="a4a8a-239">For example, "LogicalText" will target all boolean and all text-valued cells.</span></span> <span data-ttu-id="a4a8a-240">Você pode combinar dois ou três dos quatro tipos básicos.</span><span class="sxs-lookup"><span data-stu-id="a4a8a-240">You can combine any two or any three of the four basic types.</span></span> <span data-ttu-id="a4a8a-241">Os nomes desses valores de enumeração que combinam os tipos básicos sempre estão em ordem alfabética.</span><span class="sxs-lookup"><span data-stu-id="a4a8a-241">The names of these enum values that combine basic types are always in alphabetical order.</span></span> <span data-ttu-id="a4a8a-242">Portanto para combinar células com valor de erro, valor de texto e valor booleano, use "ErrorLogicalText", e não "LogicalErrorText" ou "TextErrorLogical".</span><span class="sxs-lookup"><span data-stu-id="a4a8a-242">So to combine error-valued, text-valued, and boolean-valued cells, use "ErrorLogicalText", not "LogicalErrorText" or "TextErrorLogical".</span></span> <span data-ttu-id="a4a8a-243">O parâmetro padrão "All", combina todos os quatro tipos.</span><span class="sxs-lookup"><span data-stu-id="a4a8a-243">The default parameter of "All" combines all four types.</span></span> <span data-ttu-id="a4a8a-244">O exemplo a seguir realça todas as células com fórmulas que geraram valores booleanos ou números:</span><span class="sxs-lookup"><span data-stu-id="a4a8a-244">The following example highlights all cells with formulas that produce number or boolean values:</span></span>

```js
Excel.run(function (context) {
    const sheet = context.workbook.worksheets.getActiveWorksheet();
    const usedRange = sheet.getUsedRange();
    const formulaLogicalNumberRanges = usedRange.getSpecialCells("Formulas", "LogicalNumbers");
    formulaLogicalNumberRanges.format.fill.color = "pink";

    return context.sync();
})
```

> [!NOTE]
> <span data-ttu-id="a4a8a-245">O parâmetro  `Excel.SpecialCellValueType` só pode ser usado se o parâmetro  `Excel.SpecialCellType` for "Formulas" ou "Constants".</span><span class="sxs-lookup"><span data-stu-id="a4a8a-245">The ChildObjectTypes parameter can only be used if the AccessRights parameter is set to CreateChild or DeleteChild.</span></span>

### <a name="get-rangeareas-within-rangeareas"></a><span data-ttu-id="a4a8a-246">Obter RangeAreas dentro de RangeAreas</span><span class="sxs-lookup"><span data-stu-id="a4a8a-246">Get RangeAreas within RangeAreas</span></span>

<span data-ttu-id="a4a8a-247">O tipo `RangeAreas` possui métodos  `getSpecialCells` e `getSpecialCellsOrNullObject` que obtêm os mesmos dois parâmetros.</span><span class="sxs-lookup"><span data-stu-id="a4a8a-247">The `RangeAreas` type itself also has `getSpecialCells` and `getSpecialCellsOrNullObject` methods which take the same two parameters.</span></span> <span data-ttu-id="a4a8a-248">Esses métodos retornam todas as células de destino de todos os intervalos da coleção `RangeAreas.areas` .</span><span class="sxs-lookup"><span data-stu-id="a4a8a-248">These methods return all the targeted cells from all of the ranges in the `RangeAreas.areas` collection.</span></span> <span data-ttu-id="a4a8a-249">Há uma pequena diferença no comportamento dos métodos quando chamado em um objeto `RangeAreas` , em vez de um objeto  `Range` : quando você passá "SameConditionalFormat" como o primeiro parâmetro, o método retornará todas as células que tenham a mesma formatação condicional como o célula mais à esquerda no canto superior \*do primeiro intervalo na coleção `RangeAreas.areas` \*.</span><span class="sxs-lookup"><span data-stu-id="a4a8a-249">There is one small difference in the behavior of the methods when called on a `RangeAreas` object instead of a `Range` object: when you pass "SameConditionalFormat" as the first parameter, the method returns all cells that have the same conditional formatting as the upper leftmost cell *in the first range in the `RangeAreas.areas` collection*.</span></span> <span data-ttu-id="a4a8a-250">O mesmo ponto aplica-se a "SameDataValidation": quando passados para `Range.getSpecialCells`, ele retorna todas as células que tenham a mesma regra de validação de dados como a célula mais à esquerda no canto superior *no intervalo*.</span><span class="sxs-lookup"><span data-stu-id="a4a8a-250">The same point applies to "SameDataValidation": when passed to `Range.getSpecialCells`, it returns all cells that have the same data validation rule as the upper leftmost cell *in the range*.</span></span> <span data-ttu-id="a4a8a-251">Mas, quando ele é passado para `RangeAreas.getSpecialCells`, ele retorna todas as células que possuem a mesma regra de validação de dados como a célula mais à esquerda no canto superior \*no primeiro intervalo na coleção `RangeAreas.areas`  \*.</span><span class="sxs-lookup"><span data-stu-id="a4a8a-251">But when it is passed to `RangeAreas.getSpecialCells`, it returns all cells that have the same data validation rule as the upper leftmost cell *in the first range in the `RangeAreas.areas` collection*.</span></span>

## <a name="read-properties-of-rangeareas"></a><span data-ttu-id="a4a8a-252">Propriedades de leitura das RangeAreas</span><span class="sxs-lookup"><span data-stu-id="a4a8a-252">Read properties of RangeAreas</span></span>

<span data-ttu-id="a4a8a-253">Ler os valores da propriedade de `RangeAreas` requer cuidado, pois uma determinada propriedade pode ter valores diferentes para diferentes intervalos dentro de `RangeAreas`.</span><span class="sxs-lookup"><span data-stu-id="a4a8a-253">Reading property values of `RangeAreas` requires care, because a given property may have different values for different ranges within the `RangeAreas`.</span></span> <span data-ttu-id="a4a8a-254">A regra geral é que, se um valor consistente *puder* ser retornado, ele será retornado.</span><span class="sxs-lookup"><span data-stu-id="a4a8a-254">The general rule is that if a consistent value *can* be returned it will be returned.</span></span> <span data-ttu-id="a4a8a-255">Por exemplo, no código a seguir, o código RGB para rosa (`#FFC0CB`) e `true` serão registrados no console porque ambos os intervalos no objeto  `RangeAreas` possuem um preenchimento rosa e ambos são colunas inteiras.</span><span class="sxs-lookup"><span data-stu-id="a4a8a-255">For example, in the following code, The RGB code for pink (`#FFC0CB`) and `true` will be logged to the console because both the ranges in the `RangeAreas` object have a pink fill and both are entire columns.</span></span>

```js
Excel.run(function (context) {
    const sheet = context.workbook.worksheets.getActiveWorksheet();

    // The ranges are the F column and the H column.
    const rangeAreas = sheet.getRanges("F:F, H:H");  
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

<span data-ttu-id="a4a8a-256">As coisas ficam mais complicadas quando a consistência não é possível.</span><span class="sxs-lookup"><span data-stu-id="a4a8a-256">Things get more complicated when consistency isn't possible.</span></span> <span data-ttu-id="a4a8a-257">O comportamento das propriedades `RangeAreas` segue estes três princípios:</span><span class="sxs-lookup"><span data-stu-id="a4a8a-257">The behavior of `RangeAreas` properties follows these three principles:</span></span>

- <span data-ttu-id="a4a8a-258">Uma propriedade booleana de um objeto  `RangeAreas` retorna `false` , a menos que a propriedade seja verdadeira (true) para todos os intervalos membro.</span><span class="sxs-lookup"><span data-stu-id="a4a8a-258">A boolean property of a `RangeAreas` object returns `false` unless the property is true for all the member ranges.</span></span>
- <span data-ttu-id="a4a8a-259">Propriedades não-booleanas, com exceção da propriedade `address` , retornam `null` , a menos que a propriedade correspondente em todos os intervalos membro tenha o mesmo valor.</span><span class="sxs-lookup"><span data-stu-id="a4a8a-259">Non-boolean properties, with the exception of the `address` property, return `null` unless the corresponding property on all the member ranges has the same value.</span></span>
- <span data-ttu-id="a4a8a-260">A propriedade  `address` retornará uma sequência de caracteres delimitada por vírgulas dos endereços dos intervalos membro.</span><span class="sxs-lookup"><span data-stu-id="a4a8a-260">The `address` property returns a comma-delimited string of the addresses of the member ranges.</span></span>

<span data-ttu-id="a4a8a-261">Por exemplo, o código a seguir cria um `RangeAreas` em que somente um intervalo é uma coluna inteira e apenas um é preenchido com rosa.</span><span class="sxs-lookup"><span data-stu-id="a4a8a-261">For example, the following code creates a `RangeAreas` in which only one range is an entire column and only one is filled with pink.</span></span> <span data-ttu-id="a4a8a-262">O console mostrará `null` para a cor de preenchimento, `false` para a propriedade `isEntireRow` e "Sheet1! F3:F5, Sheet1! H:H"(supondo que o nome da planilha seja "Sheet1") para a propriedade `address`.</span><span class="sxs-lookup"><span data-stu-id="a4a8a-262">The console will show `null` for the fill color, `false` for the `isEntireRow` property, and "Sheet1!F3:F5, Sheet1!H:H" (assuming the sheet name is "Sheet1") for the `address` property.</span></span> 

```js
Excel.run(function (context) {
    const sheet = context.workbook.worksheets.getActiveWorksheet();
    const rangeAreas = sheet.getRanges("F3:F5, H:H");

    const pinkColumnRange = sheet.getRange("H:H");
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

## <a name="see-also"></a><span data-ttu-id="a4a8a-263">Confira também</span><span class="sxs-lookup"><span data-stu-id="a4a8a-263">See also</span></span>

- [<span data-ttu-id="a4a8a-264">Principais conceitos da API JavaScript do Excel</span><span class="sxs-lookup"><span data-stu-id="a4a8a-264">Excel JavaScript API core concepts</span></span>](https://docs.microsoft.com/javascript/office/overview/excel-add-ins-reference-overview)
- [<span data-ttu-id="a4a8a-265">Objeto Range (API JavaScript para Excel)</span><span class="sxs-lookup"><span data-stu-id="a4a8a-265">Range Object (JavaScript API for Excel)</span></span>](https://docs.microsoft.com/javascript/api/excel/excel.range)
- <span data-ttu-id="a4a8a-266">[Objeto RangeAreas (API JavaScript para Excel)](https://docs.microsoft.com/javascript/api/excel/excel.rangeareas) (Este link pode não funcionar enquanto a API estiver em versão prévia.</span><span class="sxs-lookup"><span data-stu-id="a4a8a-266">[RangeAreas Object (JavaScript API for Excel)](https://docs.microsoft.com/javascript/api/excel/excel.rangeareas) (This link may not work while the API is in preview.</span></span> <span data-ttu-id="a4a8a-267">Como alternativa, confira [beta office.d.ts](https://appsforoffice.microsoft.com/lib/beta/hosted/office.d.ts).)</span><span class="sxs-lookup"><span data-stu-id="a4a8a-267">As an alternative, see [beta office.d.ts](https://appsforoffice.microsoft.com/lib/beta/hosted/office.d.ts).)</span></span>