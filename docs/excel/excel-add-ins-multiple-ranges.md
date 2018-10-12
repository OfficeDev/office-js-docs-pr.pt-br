---
title: Trabalhar com vários intervalos simultaneamente em suplementos do Excel
description: ''
ms.date: 9/4/2018
ms.openlocfilehash: a00bbf15b53649147fb2c2b1dfa590f15c5739be
ms.sourcegitcommit: c53f05bbd4abdfe1ee2e42fdd4f82b318b363ad7
ms.translationtype: HT
ms.contentlocale: pt-BR
ms.lasthandoff: 10/12/2018
ms.locfileid: "25506291"
---
# <a name="work-with-multiple-ranges-simultaneously-in-excel-add-ins-preview"></a><span data-ttu-id="e3bb8-102">Trabalhar com vários intervalos simultaneamente em suplementos do Excel (Versão Prévia)</span><span class="sxs-lookup"><span data-stu-id="e3bb8-102">Work with multiple ranges simultaneously in Excel add-ins (Preview)</span></span>

<span data-ttu-id="e3bb8-p101">A biblioteca JavaScript do Excel permite que o seu suplemento execute operações e defina propriedades em vários intervalos simultaneamente. Os intervalos não precisam ser contíguos. Além de simplificar o seu código, essa maneira de definir uma propriedade é mais rápida do que configurar a mesma propriedade individualmente para cada um dos intervalos.</span><span class="sxs-lookup"><span data-stu-id="e3bb8-p101">The Excel JavaScript library enables your add-in to perform operations, and set properties, on multiple ranges simultaneously. The ranges do not have to be contiguous. In addition to making your code simpler, this way of setting a property runs much faster than setting the same property individually for each of the ranges.</span></span>

> [!NOTE]
> <span data-ttu-id="e3bb8-p102">As APIs descritas neste artigo exigem a **versão 1809 do Office 2016 Clique para Executar, Build 10820.20000** ou posterior. (Talvez você precise ingressar no [programa Office Insider](https://products.office.com/office-insider) para obter o build apropriado.) Além disso, você deve carregar a versão beta da biblioteca do JavaScript do Office encontrada em [Office.js CDN](https://appsforoffice.microsoft.com/lib/beta/hosted/office.js). Por último, infelizmente ainda não temos páginas de referência para essas APIs. Mas o tipo de arquivo de definição a seguir traz descrições para eles: [office.d.ts beta](https://appsforoffice.microsoft.com/lib/beta/hosted/office.d.ts).</span><span class="sxs-lookup"><span data-stu-id="e3bb8-p102">The APIs described in this article require **Office 2016 Click-to-Run version 1809 Build 10820.20000** or later. (You may need to join the [Office Insider program](https://products.office.com/office-insider) to get an appropriate build.) Also, you must load the beta version of the Office JavaScript library from [Office.js CDN](https://appsforoffice.microsoft.com/lib/beta/hosted/office.js). Finally, we don't have reference pages for these APIs yet. But the following definition type file has descriptions for them: [beta office.d.ts](https://appsforoffice.microsoft.com/lib/beta/hosted/office.d.ts).</span></span>

## <a name="rangeareas"></a><span data-ttu-id="e3bb8-110">RangeAreas</span><span class="sxs-lookup"><span data-stu-id="e3bb8-110">RangeAreas</span></span>

<span data-ttu-id="e3bb8-p103">Um conjunto de intervalos (possivelmente não contíguos) é representado por um objeto `Excel.RangeAreas`. Ele tem propriedades e métodos semelhantes ao tipo `Range` (vários com nomes semelhantes ou iguais), mas alguns ajustes foram feitos em:</span><span class="sxs-lookup"><span data-stu-id="e3bb8-p103">A set of (possibly discontiguous) ranges is represented by an `Excel.RangeAreas` object. It has properties and methods similar to the `Range` type (many with the same, or similar, names), but adjustments have been made to:</span></span>

- <span data-ttu-id="e3bb8-113">Os tipos de dados para as propriedades e o comportamento dos setters e getters.</span><span class="sxs-lookup"><span data-stu-id="e3bb8-113">The data types for properties and the behavior of the setters and getters.</span></span>
- <span data-ttu-id="e3bb8-114">Tipos de dados dos parâmetros do método e nos comportamentos do método.</span><span class="sxs-lookup"><span data-stu-id="e3bb8-114">The data types of method parameters and the method behaviors.</span></span>
- <span data-ttu-id="e3bb8-115">Valores retornados dos tipos de dados do método.</span><span class="sxs-lookup"><span data-stu-id="e3bb8-115">The data types of method return values.</span></span>

<span data-ttu-id="e3bb8-116">Alguns exemplos:</span><span class="sxs-lookup"><span data-stu-id="e3bb8-116">Some examples:</span></span>

- <span data-ttu-id="e3bb8-117">`RangeAreas` tem uma propriedade  `address` que retorna uma sequência de caracteres delimitada por vírgula do intervalo de endereços, em vez de apenas um endereço como na propriedade `Range.address` .</span><span class="sxs-lookup"><span data-stu-id="e3bb8-117">`RangeAreas` has an `address` property that returns a comma-delimited string of range addresses, instead of just one address as with the `Range.address` property.</span></span>
- <span data-ttu-id="e3bb8-p104">`RangeAreas` tem uma propriedade `dataValidation` que retorna um objeto `DataValidation` que representa a validação de dados de todos os intervalos em `RangeAreas`, caso seja consistente. A propriedade é `null` se objetos `DataValidation` idênticos não forem aplicados a todos os intervalos em `RangeAreas`. Esse é um princípio geral, mas não universal, do objeto `RangeAreas`: *Se uma propriedade não tiver valores consistentes em todos os intervalos em `RangeAreas`, então, é `null`.* Consulte [Propriedades de leitura de RangeAreas](#reading-properties-of-rangeareas) para obter mais informações e algumas exceções.</span><span class="sxs-lookup"><span data-stu-id="e3bb8-p104">`RangeAreas` has a `dataValidation` property that returns a `DataValidation` object that represents the data validation of all the ranges in the `RangeAreas`, if it is consistent. The property is `null` if identical `DataValidation` objects are not applied to all the all the ranges in the `RangeAreas`. This is a general, but not universal, principle with the `RangeAreas` object: *If a property does not have consistent values on all the all the ranges in the `RangeAreas`, then it is `null`.* See [Reading properties of RangeAreas](#reading-properties-of-rangeareas) for more information and some exceptions.</span></span>
- <span data-ttu-id="e3bb8-122">`RangeAreas.cellCount` obtém o número total de células em todos os intervalos de `RangeAreas`.</span><span class="sxs-lookup"><span data-stu-id="e3bb8-122">`RangeAreas.cellCount` gets the total number of cells in all the ranges in the `RangeAreas`.</span></span>
- <span data-ttu-id="e3bb8-123">`RangeAreas.calculate` recalcula as células de todos os intervalos de `RangeAreas`.</span><span class="sxs-lookup"><span data-stu-id="e3bb8-123">`RangeAreas.calculate` recalculates the cells of all the ranges in the `RangeAreas`.</span></span>
- <span data-ttu-id="e3bb8-p105">`RangeAreas.getEntireColumn` e `RangeAreas.getEntireRow` retorna outro objeto `RangeAreas` que representa todas as colunas (ou linhas) em todos os intervalos de `RangeAreas`. Por exemplo, se `RangeAreas` representa "A1:C4" e "F14:L15", então, `RangeAreas.getEntireColumn` retorna um objeto `RangeAreas` que representa "A:C" e "F:L".</span><span class="sxs-lookup"><span data-stu-id="e3bb8-p105">`RangeAreas.getEntireColumn` and `RangeAreas.getEntireRow` return another `RangeAreas` object that represents all of the columns (or rows) in all the ranges in the `RangeAreas`. For example, if the `RangeAreas` represents "A1:C4" and "F14:L15", then `RangeAreas.getEntireColumn` returns a `RangeAreas` object that represents "A:C" and "F:L".</span></span>
- <span data-ttu-id="e3bb8-126">`RangeAreas.copyFrom` pode receber um parâmetro `Range` ou `RangeAreas` que representa o(s) intervalo(s) de origem da operação de cópia.</span><span class="sxs-lookup"><span data-stu-id="e3bb8-126">`RangeAreas.copyFrom` can take either a `Range` or a `RangeAreas` parameter representing the source range(s) of the copy operation.</span></span>

#### <a name="complete-list-of-range-members-that-are-also-available-on-rangeareas"></a><span data-ttu-id="e3bb8-127">Lista completa dos membros de Range que também estão disponíveis em RangeAreas</span><span class="sxs-lookup"><span data-stu-id="e3bb8-127">Complete list of Range members that are also available on RangeAreas</span></span>

##### <a name="properties"></a><span data-ttu-id="e3bb8-128">Propriedades</span><span class="sxs-lookup"><span data-stu-id="e3bb8-128">Properties</span></span>

<span data-ttu-id="e3bb8-p106">Conheça as [Propriedades de leitura de RangeAreas](#reading-properties-of-rangeareas) antes de escrever códigos que leiam as propriedades listadas. Há sutilezas quanto a o que é retornado.</span><span class="sxs-lookup"><span data-stu-id="e3bb8-p106">Be familiar with [Reading properties of RangeAreas](#reading-properties-of-rangeareas) before you write code that reads any properties listed. There are subtleties to what gets returned.</span></span>

- <span data-ttu-id="e3bb8-131">address</span><span class="sxs-lookup"><span data-stu-id="e3bb8-131">address</span></span>
- <span data-ttu-id="e3bb8-132">addressLocal</span><span class="sxs-lookup"><span data-stu-id="e3bb8-132">addressLocal</span></span>
- <span data-ttu-id="e3bb8-133">cellCount</span><span class="sxs-lookup"><span data-stu-id="e3bb8-133">cellCount</span></span>
- <span data-ttu-id="e3bb8-134">conditionalFormats</span><span class="sxs-lookup"><span data-stu-id="e3bb8-134">conditionalFormats</span></span>
- <span data-ttu-id="e3bb8-135">context</span><span class="sxs-lookup"><span data-stu-id="e3bb8-135">context</span></span>
- <span data-ttu-id="e3bb8-136">dataValidation</span><span class="sxs-lookup"><span data-stu-id="e3bb8-136">dataValidation</span></span>
- <span data-ttu-id="e3bb8-137">format</span><span class="sxs-lookup"><span data-stu-id="e3bb8-137">format</span></span>
- <span data-ttu-id="e3bb8-138">isEntireColumn</span><span class="sxs-lookup"><span data-stu-id="e3bb8-138">isEntireColumn</span></span>
- <span data-ttu-id="e3bb8-139">isEntireRow</span><span class="sxs-lookup"><span data-stu-id="e3bb8-139">isEntireRow</span></span>
- <span data-ttu-id="e3bb8-140">style</span><span class="sxs-lookup"><span data-stu-id="e3bb8-140">style</span></span>
- <span data-ttu-id="e3bb8-141">worksheet</span><span class="sxs-lookup"><span data-stu-id="e3bb8-141">worksheet</span></span>

##### <a name="methods"></a><span data-ttu-id="e3bb8-142">Métodos</span><span class="sxs-lookup"><span data-stu-id="e3bb8-142">Methods</span></span>

<span data-ttu-id="e3bb8-143">Métodos de Range em versão prévia estão marcados.</span><span class="sxs-lookup"><span data-stu-id="e3bb8-143">Range methods in preview are marked.</span></span>

- <span data-ttu-id="e3bb8-144">calculate()</span><span class="sxs-lookup"><span data-stu-id="e3bb8-144">calculate()</span></span>
- <span data-ttu-id="e3bb8-145">clear()</span><span class="sxs-lookup"><span data-stu-id="e3bb8-145">clear()</span></span>
- <span data-ttu-id="e3bb8-146">convertDataTypeToText() (versão prévia)</span><span class="sxs-lookup"><span data-stu-id="e3bb8-146">convertDataTypeToText() (preview)</span></span>
- <span data-ttu-id="e3bb8-147">convertToLinkedDataType() (versão prévia)</span><span class="sxs-lookup"><span data-stu-id="e3bb8-147">convertToLinkedDataType() (preview)</span></span>
- <span data-ttu-id="e3bb8-148">copyFrom() (versão prévia)</span><span class="sxs-lookup"><span data-stu-id="e3bb8-148">copyFrom() (preview)</span></span>
- <span data-ttu-id="e3bb8-149">getEntireColumn()</span><span class="sxs-lookup"><span data-stu-id="e3bb8-149">getEntireColumn()</span></span>
- <span data-ttu-id="e3bb8-150">getEntireRow()</span><span class="sxs-lookup"><span data-stu-id="e3bb8-150">getEntireRow()</span></span>
- <span data-ttu-id="e3bb8-151">getIntersection()</span><span class="sxs-lookup"><span data-stu-id="e3bb8-151">getIntersection()</span></span>
- <span data-ttu-id="e3bb8-152">getIntersectionOrNullObject()</span><span class="sxs-lookup"><span data-stu-id="e3bb8-152">getIntersectionOrNullObject()</span></span>
- <span data-ttu-id="e3bb8-153">getOffsetRange() (chamado getOffsetRangeAreas no objeto RangeAreas)</span><span class="sxs-lookup"><span data-stu-id="e3bb8-153">getOffsetRange() (named getOffsetRangeAreas on the RangeAreas object)</span></span>
- <span data-ttu-id="e3bb8-154">getSpecialCells() (versão prévia)</span><span class="sxs-lookup"><span data-stu-id="e3bb8-154">getSpecialCells() (preview)</span></span>
- <span data-ttu-id="e3bb8-155">getSpecialCellsOrNullObject() (versão prévia)</span><span class="sxs-lookup"><span data-stu-id="e3bb8-155">getSpecialCellsOrNullObject() (preview)</span></span>
- <span data-ttu-id="e3bb8-156">getTables() (versão prévia)</span><span class="sxs-lookup"><span data-stu-id="e3bb8-156">getTables() (preview)</span></span>
- <span data-ttu-id="e3bb8-157">getUsedRange() (chamado getUsedRangeAreas no objeto RangeAreas)</span><span class="sxs-lookup"><span data-stu-id="e3bb8-157">getUsedRange() (named getUsedRangeAreas on the RangeAreas object)</span></span>
- <span data-ttu-id="e3bb8-158">getUsedRangeOrNullObject() (chamado getUsedRangeAreasOrNullObject no objeto RangeAreas)</span><span class="sxs-lookup"><span data-stu-id="e3bb8-158">getUsedRangeOrNullObject() (named getUsedRangeAreasOrNullObject on the RangeAreas object)</span></span>
- <span data-ttu-id="e3bb8-159">load()</span><span class="sxs-lookup"><span data-stu-id="e3bb8-159">load()</span></span>
- <span data-ttu-id="e3bb8-160">set()</span><span class="sxs-lookup"><span data-stu-id="e3bb8-160">set\*</span></span>
- <span data-ttu-id="e3bb8-161">setDirty() (versão prévia)</span><span class="sxs-lookup"><span data-stu-id="e3bb8-161">setDirty() (preview)</span></span>
- <span data-ttu-id="e3bb8-162">toJSON()</span><span class="sxs-lookup"><span data-stu-id="e3bb8-162">toJSON()</span></span>
- <span data-ttu-id="e3bb8-163">track()</span><span class="sxs-lookup"><span data-stu-id="e3bb8-163">track</span></span>
- <span data-ttu-id="e3bb8-164">untrack()</span><span class="sxs-lookup"><span data-stu-id="e3bb8-164">untrack()</span></span>

### <a name="rangearea-specific-properties-and-methods"></a><span data-ttu-id="e3bb8-165">Propriedades e métodos específicos de RangeArea</span><span class="sxs-lookup"><span data-stu-id="e3bb8-165">RangeArea-specific properties and methods</span></span>

<span data-ttu-id="e3bb8-p107">O tipo `RangeAreas` tem algumas propriedades e métodos que não estão no objeto `Range`. Veja a seguir uma seleção deles:</span><span class="sxs-lookup"><span data-stu-id="e3bb8-p107">The `RangeAreas` type has some properties and methods that are not on the `Range` object. The following is a selection of them:</span></span>

- <span data-ttu-id="e3bb8-p108">`areas`: Um objeto `RangeCollection` que contém todos os intervalos representados pelo objeto `RangeAreas`. O objeto `RangeCollection` também é novo e é semelhante a outros objetos da coleção do Excel. Ele tem uma propriedade `items` que é uma matriz de objetos `Range` que representam os intervalos.</span><span class="sxs-lookup"><span data-stu-id="e3bb8-p108">`areas`: A `RangeCollection` object that contains all of the ranges represented by the `RangeAreas` object. The `RangeCollection` object is also new and is similar to other Excel collection objects. It has an `items` property which is an array of `Range` objects representing the ranges.</span></span>
- <span data-ttu-id="e3bb8-171">`areaCount`: O número total de intervalos em `RangeAreas`.</span><span class="sxs-lookup"><span data-stu-id="e3bb8-171">`areaCount`: The total number of ranges in the `RangeAreas`.</span></span>
- <span data-ttu-id="e3bb8-172">`getOffsetRangeAreas`: Funciona exatamente como [Range.getOffsetRange](https://docs.microsoft.com/javascript/api/excel/excel.range#getoffsetrange-rowoffset--columnoffset-), exceto que um `RangeAreas` é retornado e contém intervalos que são um deslocamento de um dos intervalos no `RangeAreas` original.</span><span class="sxs-lookup"><span data-stu-id="e3bb8-172">`getOffsetRangeAreas`: Works just like [Range.getOffsetRange](https://docs.microsoft.com/javascript/api/excel/excel.range#getoffsetrange-rowoffset--columnoffset-), except that a `RangeAreas` is returned and it contains ranges that are each offset from one of the ranges in the original `RangeAreas`.</span></span>

## <a name="create-rangeareas-and-set-properties"></a><span data-ttu-id="e3bb8-173">Criar RangeAreas e definir propriedades</span><span class="sxs-lookup"><span data-stu-id="e3bb8-173">Create RangeAreas and set properties</span></span>

<span data-ttu-id="e3bb8-174">Você pode criar o objeto  `RangeAreas` de duas formas básicas:</span><span class="sxs-lookup"><span data-stu-id="e3bb8-174">You can create `RangeAreas` object in two basic ways:</span></span>

- <span data-ttu-id="e3bb8-p109">Chame `Worksheet.getRanges()` e passe para ele uma sequência de caracteres com endereços de intervalo delimitados por vírgula. Se algum dos intervalos que você deseja incluir tiver sido transformado em um [getNamedItem](https://docs.microsoft.com/javascript/api/excel/excel.nameditem), você pode incluir o nome, em vez do endereço, na sequência de caracteres.</span><span class="sxs-lookup"><span data-stu-id="e3bb8-p109">Call `Worksheet.getRanges()` and pass it a string with comma-delimited range addresses. If any range you want to include has been made into a [NamedItem](https://docs.microsoft.com/javascript/api/excel/excel.nameditem), you can include the name, instead of the address, in the string.</span></span>
- <span data-ttu-id="e3bb8-p110">Chame `Workbook.getSelectedRanges()`. Esse método retorna `RangeAreas` que representa todos os intervalos selecionados na planilha ativa no momento.</span><span class="sxs-lookup"><span data-stu-id="e3bb8-p110">Call `Workbook.getSelectedRanges()`. This method returns a `RangeAreas` representing all the ranges that are selected on the currently active worksheet.</span></span>

<span data-ttu-id="e3bb8-179">Depois que você tiver um objeto `RangeAreas` , você pode criar outros usando os métodos no objeto que retornam `RangeAreas`, como `getOffsetRangeAreas` e `getIntersection`.</span><span class="sxs-lookup"><span data-stu-id="e3bb8-179">Once you have a `RangeAreas` object, you can create others using the methods on the object that return `RangeAreas` such as `getOffsetRangeAreas` and `getIntersection`.</span></span>

> [!NOTE]
> <span data-ttu-id="e3bb8-p111">Você não pode adicionar intervalos adicionais diretamente em um objeto `RangeAreas`. Por exemplo, a coleção em `RangeAreas.areas` não tem um método `add`.</span><span class="sxs-lookup"><span data-stu-id="e3bb8-p111">You cannot directly add additional ranges to a `RangeAreas` object. For example, the collection in `RangeAreas.areas` does not have an `add` method.</span></span>


> [!WARNING] 
> <span data-ttu-id="e3bb8-p112">Não tente adicionar ou excluir membros da matriz `RangeAreas.areas.items` diretamente. Isso causará um comportamento indesejável no seu código. Por exemplo, é possível inserir um objeto `Range` adicional na matriz, mas isso irá causar erros porque métodos e propriedades `RangeAreas` se comportam como se o novo item não estivesse lá. Por exemplo, a propriedade `areaCount` não inclui intervalos inseridos dessa maneira e `RangeAreas.getItemAt(index)` gera um erro se `index` for maior do que `areasCount-1`. Da mesma forma, excluir um objeto `Range` da matriz `RangeAreas.areas.items` obtendo uma referência a ele e chamando o método `Range.delete` causará erros: embora o objeto `Range` *seja* excluído, as propriedades e métodos do objeto `RangeAreas` pai se comportam, ou tentam se comportar, como se ele ainda existisse. Por exemplo, se o seu código chamar `RangeAreas.calculate`, o Office tentará calcular o intervalo, mas apresentará um erro, pois o objeto range não existe mais.</span><span class="sxs-lookup"><span data-stu-id="e3bb8-p112">Do not attempt to directly add or delete members of the the `RangeAreas.areas.items` array. This will lead to undesirable behavior in your code. For example, it is possible to push an additional `Range` object onto the array, but doing so will cause errors because `RangeAreas` properties and methods behave as if the new item isn't there. For example, the `areaCount` property does not include ranges pushed in this way, and the `RangeAreas.getItemAt(index)` throws an error if `index` is larger than `areasCount-1`. Similarly, deleting a `Range` object in the `RangeAreas.areas.items` array by getting a reference to it and calling its `Range.delete` method causes bugs: although the `Range` object *is* deleted, the properties and methods of the parent `RangeAreas` object behave, or try to, as if it is still in existence. For example, if your code calls `RangeAreas.calculate`, Office will try to calculate the range, but will error because the range object is gone.</span></span>

<span data-ttu-id="e3bb8-188">Configurar uma propriedade em um `RangeAreas` define a propriedade correspondente em todos os intervalos na coleção `RangeAreas.areas` .</span><span class="sxs-lookup"><span data-stu-id="e3bb8-188">Setting a property on a `RangeAreas` sets the corresponding property on all the ranges in the `RangeAreas.areas` collection.</span></span>

<span data-ttu-id="e3bb8-p113">A seguir, veja um exemplo de definição de uma propriedade em vários intervalos. A função realça os intervalos **F3:F5** e **H3:H5**.</span><span class="sxs-lookup"><span data-stu-id="e3bb8-p113">The following is an example of setting a property on multiple ranges. The function highlights the ranges **F3:F5** and **H3:H5**.</span></span>

```js
Excel.run(function (context) {
    const sheet = context.workbook.worksheets.getActiveWorksheet();
    const rangeAreas = sheet.getRanges("F3:F5, H3:H5");
    rangeAreas.format.fill.color = "pink";

    return context.sync();
})
```

<span data-ttu-id="e3bb8-p114">Este exemplo se aplica a cenários nos quais você pode codificar os endereços do intervalo que você passa para `getRanges` ou facilmente calculá-los no tempo de execução. Alguns dos cenários em que isso seria possível incluem:</span><span class="sxs-lookup"><span data-stu-id="e3bb8-p114">This example applies to scenarios in which you can hard code the range addresses that you pass to `getRanges` or easily calculate them at runtime. Some of the scenarios in which this would be true include:</span></span> 

- <span data-ttu-id="e3bb8-193">O código é executado no contexto de um modelo conhecido.</span><span class="sxs-lookup"><span data-stu-id="e3bb8-193">The code runs in the context of a known template.</span></span>
- <span data-ttu-id="e3bb8-194">O código é executado no contexto de dados importados onde o esquema dos dados é conhecido.</span><span class="sxs-lookup"><span data-stu-id="e3bb8-194">The code runs in the context of imported data where the schema of the data is known.</span></span>

<span data-ttu-id="e3bb8-p115">Quando você não sabe durante a codificação quais os intervalos em que você precisa para operar, você deve descobri-los no tempo de execução. A próxima seção discute esses cenários.</span><span class="sxs-lookup"><span data-stu-id="e3bb8-p115">When you can't know at coding-time which ranges you need to operate on, you must discover them at runtime. The next section discusses these scenarios.</span></span>

### <a name="discover-range-areas-programmatically"></a><span data-ttu-id="e3bb8-197">Descobrir áreas de intervalo programaticamente</span><span class="sxs-lookup"><span data-stu-id="e3bb8-197">Discover range areas programmatically</span></span>

<span data-ttu-id="e3bb8-p116">Os métodos `Range.getSpecialCells()` e `Range.getSpecialCellsOrNullObject()` permitem que você descubra, durante o tempo de execução, os intervalos em que você deseja operar com base nas características das células e no tipo dos valores das células. Aqui estão as assinaturas dos métodos do arquivo de tipos de dados TypeScript:</span><span class="sxs-lookup"><span data-stu-id="e3bb8-p116">The `Range.getSpecialCells()` and `Range.getSpecialCellsOrNullObject()` methods enable you to find at runtime the ranges that you want to operate on based on the characteristics of the cells and the type of the values of the cells. Here are the signatures of the methods from the TypeScript data types file:</span></span>

```typescript
getSpecialCells(cellType: Excel.SpecialCellType, cellValueType?: Excel.SpecialCellValueType): Excel.RangeAreas;
```

```typescript
getSpecialCellsOrNullObject(cellType: Excel.SpecialCellType, cellValueType?: Excel.SpecialCellValueType): Excel.RangeAreas;
```

<span data-ttu-id="e3bb8-p117">A seguir, veja um exemplo de uso do primeiro. Sobre este código, observe:</span><span class="sxs-lookup"><span data-stu-id="e3bb8-p117">The following is an example of using the first one. About this code, note:</span></span>

- <span data-ttu-id="e3bb8-202">Ele limita a parte da planilha que precisa ser pesquisada chamando primeiro `Worksheet.getUsedRange` e chamando `getSpecialCells` somente para aquele intervalo.</span><span class="sxs-lookup"><span data-stu-id="e3bb8-202">It limits the part of the sheet that needs to be searched by first calling `Worksheet.getUsedRange` and calling `getSpecialCells` for only that range.</span></span>
- <span data-ttu-id="e3bb8-p118">Ele passa como um parâmetro para `getSpecialCells` a versão de sequência de caracteres de um valor a partir da enumeração `Excel.SpecialCellType`. Alguns dos outros valores que podem ser passados, em vez disso, são "Blanks" para células vazias, "Constants" para células com valores literais em vez de fórmulas e "SameConditionalFormat" para células com a mesma formatação condicional que a primeira célula em `usedRange`. A primeira célula é a célula superior mais à esquerda. Para obter uma lista completa dos valores na enumeração, consulte [office.d.ts beta](https://appsforoffice.microsoft.com/lib/beta/hosted/office.d.ts).</span><span class="sxs-lookup"><span data-stu-id="e3bb8-p118">It passes as a parameter to `getSpecialCells` the string version of a value from the `Excel.SpecialCellType` enum. Some of the other values that could be passed instead are "Blanks" for empty cells, "Constants" for cells with literal values instead of formulas, and "SameConditionalFormat" for cells that have the same conditional formatting as the first cell in the `usedRange`. The first cell is the upper leftmost cell. For a complete list of the values in the enum, see [beta office.d.ts](https://appsforoffice.microsoft.com/lib/beta/hosted/office.d.ts).</span></span>
- <span data-ttu-id="e3bb8-207">O método `getSpecialCells` retorna um objeto `RangeAreas`, portanto todas as células com fórmulas serão cor de rosa, mesmo que não sejam contíguas.</span><span class="sxs-lookup"><span data-stu-id="e3bb8-207">The `getSpecialCells` method returns a `RangeAreas` object, so all of the cells with formulas will be colored pink even if they are not all contiguous.</span></span> 

```js
Excel.run(function (context) {
    const sheet = context.workbook.worksheets.getActiveWorksheet();
    const usedRange = sheet.getUsedRange();
    const formulaRanges = usedRange.getSpecialCells("Formulas");
    formulaRanges.format.fill.color = "pink";

    return context.sync();
})
```

<span data-ttu-id="e3bb8-p119">Em alguns casos, o intervalo não tem *nenhuma* célula com a característica alvo. Se `getSpecialCells` não encontrar nenhuma, ele gera um erro de **ItemNotFound** . Isso desviaria o fluxo de controle para um bloco/método `catch`, casa haja um. Se não houver, o erro interrompe a função. Pode haver cenários nos quais emitir o erro é exatamente o que você deseja que aconteça, quando não há nenhuma célula com a característica alvo.</span><span class="sxs-lookup"><span data-stu-id="e3bb8-p119">Sometimes the range doesn't have *any* cells with the targeted characteristic. If `getSpecialCells` doesn't find any, it throws an **ItemNotFound** error. This would divert the flow of control to a `catch` block/method, if there is one. If there isn't, the error halts the function. There may be scenarios in which throwing the error is exactly what you want to happen when there are no cells with the targeted characteristic.</span></span> 

<span data-ttu-id="e3bb8-p120">Mas há cenários nos quais é normal, mas talvez incomum, que não haja nenhuma célula correspondente; seu código deve verificar essa possibilidade e lidar com ela sem dificuldades e sem gerar um erro. Para esses cenários, use o método `getSpecialCellsOrNullObject` e teste a propriedade `RangeAreas.isNullObject`. Veja um exemplo a seguir. Nota sobre este código:</span><span class="sxs-lookup"><span data-stu-id="e3bb8-p120">But in scenarios in which it is normal, but perhaps uncommon, for there to be no matching cells; your code should check for this possibility and handle it gracefully without throwing an error. For these scenarios, use the `getSpecialCellsOrNullObject` method and test the `RangeAreas.isNullObject` property. The following is an example. Note about this code:</span></span>

- <span data-ttu-id="e3bb8-p121">O método `getSpecialCellsOrNullObject` sempre retorna um objeto proxy, isso significa que nunca é `null` no sentido comum do JavaScript. Mas se nenhuma célula correspondente for encontrada, a propriedade `isNullObject` do objeto é definida como `true`.</span><span class="sxs-lookup"><span data-stu-id="e3bb8-p121">The `getSpecialCellsOrNullObject` method always returns a proxy object, so it is never `null` in the ordinary JavaScript sense. But if no matching cells are found, the `isNullObject` property of the object is set to `true`.</span></span>
- <span data-ttu-id="e3bb8-p122">Ele chama `context.sync` *antes* de testar a propriedade `isNullObject`. Esse é um requisito de todos os métodos e propriedades `*OrNullObject`, pois você sempre precisa carregar e sincronizar uma propriedade para poder lê-la. No entanto, não é necessário carregar *explicitamente* a propriedade `isNullObject`. Ela é carregado automaticamente por `context.sync` , mesmo que `load` não seja chamado no objeto. Para obter mais informações, consulte [\*OrNullObject](https://docs.microsoft.com/office/dev/add-ins/excel/excel-add-ins-advanced-concepts#42ornullobject-methods).</span><span class="sxs-lookup"><span data-stu-id="e3bb8-p122">It calls `context.sync` *before* it tests the `isNullObject` property. This is a requirement with all `*OrNullObject` methods and properties, because you always have to load and sync a property in order to read it. However, it is not necessary to *explicitly* load the `isNullObject` property. It is automatically loaded by the `context.sync` even if `load` is not called on the object. For more information, see [\*OrNullObject](https://docs.microsoft.com/office/dev/add-ins/excel/excel-add-ins-advanced-concepts#42ornullobject-methods).</span></span>
- <span data-ttu-id="e3bb8-p123">Para testar esse código, selecione um intervalo que não tenha células com fórmulas e execute-o. Depois, selecione um intervalo que tenha pelo menos uma célula com fórmula e execute-o novamente.</span><span class="sxs-lookup"><span data-stu-id="e3bb8-p123">You can test this code by first selecting a range that has no formula cells and running it. Then select a range that has at least one cell with a formula and run it again.</span></span>

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

<span data-ttu-id="e3bb8-226">Para manter a simplicidade, todos os outros exemplos neste artigo usam o método `getSpecialCells` em vez de `getSpecialCellsOrNullObject`.</span><span class="sxs-lookup"><span data-stu-id="e3bb8-226">For simplicity, all other examples in this article use the `getSpecialCells` method instead of  `getSpecialCellsOrNullObject`.</span></span>

#### <a name="narrow-the-target-cells-with-cell-value-types"></a><span data-ttu-id="e3bb8-227">Restringir as células de destino com tipos de valores de célula</span><span class="sxs-lookup"><span data-stu-id="e3bb8-227">Narrow the target cells with cell value types</span></span>

<span data-ttu-id="e3bb8-p124">Este é um segundo parâmetro opcional, de `Excel.SpecialCellValueType` tipo enumerado, que restringe ainda mais as células alvo. Você pode usá-lo somente quando passa "Formulas" ou "Constants" para `getSpecialCells` ou `getSpecialCellsOrNullObject`. O parâmetro especifica que você deseja somente células com certos tipos de valores. Existem quatro tipos básicos: "Error", "Logical" (que significa booleano), "Numbers", e "Text". (A enumeração tem outros valores além desses quatro discutidos adiante.) Veja um exemplo a seguir. Sobre este código, note:</span><span class="sxs-lookup"><span data-stu-id="e3bb8-p124">There is an optional second parameter, of enum type `Excel.SpecialCellValueType`, that further narrows down the cells to target. You can use it only when you pass either "Formulas" or "Constants" to `getSpecialCells` or `getSpecialCellsOrNullObject`. The parameter specifies that you only want cells with certain types of values. There are four basic types: "Error", "Logical" (which means boolean), "Numbers", and "Text". (The enum has other values besides these four which are discussed below.) The following is an example. About this code, note:</span></span>

- <span data-ttu-id="e3bb8-p125">Ele realçará somente células que tenha um valor de número literal. Ele não realçará células que tenham uma fórmula (mesmo que o resultado seja um número), um valor booleano, texto ou células de estado de erro.</span><span class="sxs-lookup"><span data-stu-id="e3bb8-p125">It will only highlight cells that have a literal number value. It will not highlight cells that have a formula (even if the result is a number) or a boolean, text, or error state cells.</span></span>
- <span data-ttu-id="e3bb8-236">Para testar o código, certifique-se de que a planilha possui algumas células com valores numéricos literais, outras com outros tipos de valores literais e algumas células com fórmulas.</span><span class="sxs-lookup"><span data-stu-id="e3bb8-236">To test the code, be sure the worksheet has some cells with literal number values, some with other kinds of literal values, and some with formulas.</span></span>

```js
Excel.run(function (context) {
    const sheet = context.workbook.worksheets.getActiveWorksheet();
    const usedRange = sheet.getUsedRange();
    const constantNumberRanges = usedRange.getSpecialCells("Constants", "Numbers");
    constantNumberRanges.format.fill.color = "pink";

    return context.sync();
})
```

<span data-ttu-id="e3bb8-p126">Às vezes, é necessário operar em mais de um tipo de valor de célula, como células com valores todos de texto ou todos booleanos ("Logical"). A enumeração `Excel.SpecialCellValueType` possui valores que permitem que você combine tipos. Por exemplo, "LogicalText" tem como alvo todas as células com valores completamente de texto ou completamente booleanos. Você pode combinar quaisquer dois ou três dos quatro tipos básicos. Os nomes desses valores enumerados que combinam tipos básicos sempre seguem a ordem alfabética. Então, para combinar células com valores de erros, texto e booleanos, use "ErrorLogicalText", não "LogicalErrorText" nem "TextErrorLogical". O parâmetro padrão "all" combina todos os quatro tipos. O exemplo a seguir realça todas as células com fórmulas que produzem valores numéricos ou booleanos.</span><span class="sxs-lookup"><span data-stu-id="e3bb8-p126">Sometimes you need to operate on more than one cell value type, such as all text-valued and all boolean-valued ("Logical") cells. The `Excel.SpecialCellValueType` enum has values that let you combine types. For example, "LogicalText" will target all boolean and all text-valued cells. You can combine any two or any three of the four basic types. The names of these enum values that combine basic types are always in alphabetical order. So to combine error-valued, text-valued, and boolean-valued cells, use "ErrorLogicalText", not "LogicalErrorText" or "TextErrorLogical". The default parameter of "All" combines all four types. The following example highlights all cells with formulas that produce number or boolean values:</span></span>

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
> <span data-ttu-id="e3bb8-245">O parâmetro  `Excel.SpecialCellValueType` só pode ser usado se o parâmetro  `Excel.SpecialCellType` for "Formulas" ou "Constants".</span><span class="sxs-lookup"><span data-stu-id="e3bb8-245">The ChildObjectTypes parameter can only be used if the AccessRights parameter is set to CreateChild or DeleteChild.</span></span>

### <a name="get-rangeareas-within-rangeareas"></a><span data-ttu-id="e3bb8-246">Obter RangeAreas dentro de RangeAreas</span><span class="sxs-lookup"><span data-stu-id="e3bb8-246">Get RangeAreas within RangeAreas</span></span>

<span data-ttu-id="e3bb8-p127">O próprio tipo `RangeAreas` também tem métodos `getSpecialCells` e `getSpecialCellsOrNullObject` que usam os mesmos dois parâmetros. Esses métodos retornam todas as células alvo de todos os intervalos do conjunto `RangeAreas.areas`. Há uma pequena diferença no comportamento dos métodos quando chamados em um objeto `RangeAreas`, em vez de um objeto `Range`: quando você passa "SameConditionalFormat" como o primeiro parâmetro, o método retorna todas as células com formatação condicional igual a da célula superior mais à esquerda *do primeiro intervalo da `RangeAreas.areas` coleção*. O mesmo aplica-se a "SameDataValidation": quando passado para `Range.getSpecialCells`, retorna todas as células com a mesma regra de validação de dados que a célula superior mais à esquerda *no intervalo*. Mas quando ele é passado para `RangeAreas.getSpecialCells`, retorna todas as células com a mesma regra de validação de dados que a célula superior mais à esquerda *do primeiro intervalo da `RangeAreas.areas` coleção*.</span><span class="sxs-lookup"><span data-stu-id="e3bb8-p127">The `RangeAreas` type itself also has `getSpecialCells` and `getSpecialCellsOrNullObject` methods which take the same two parameters. These methods return all the targeted cells from all of the ranges in the `RangeAreas.areas` collection. There is one small difference in the behavior of the methods when called on a `RangeAreas` object instead of a `Range` object: when you pass "SameConditionalFormat" as the first parameter, the method returns all cells that have the same conditional formatting as the upper leftmost cell *in the first range in the `RangeAreas.areas` collection*. The same point applies to "SameDataValidation": when passed to `Range.getSpecialCells`, it returns all cells that have the same data validation rule as the upper leftmost cell *in the range*. But when it is passed to `RangeAreas.getSpecialCells`, it returns all cells that have the same data validation rule as the upper leftmost cell *in the first range in the `RangeAreas.areas` collection*.</span></span>

## <a name="read-properties-of-rangeareas"></a><span data-ttu-id="e3bb8-252">Propriedades de leitura das RangeAreas</span><span class="sxs-lookup"><span data-stu-id="e3bb8-252">Read properties of RangeAreas</span></span>

<span data-ttu-id="e3bb8-p128">A leitura de valores de propriedade de `RangeAreas` requer cuidado, pois uma determinada propriedade pode ter valores diferentes para intervalos diferentes dentro de `RangeAreas`. A regra geral é que, se um valor consistente *pode* ser retornado, ele será retornado. Por exemplo, no código a seguir, o código RGB para rosa (`#FFC0CB`) e `true` serão registrados no console pois ambos os intervalos no objeto `RangeAreas` possuem preenchimento rosa e ambos são colunas inteiras.</span><span class="sxs-lookup"><span data-stu-id="e3bb8-p128">Reading property values of `RangeAreas` requires care, because a given property may have different values for different ranges within the `RangeAreas`. The general rule is that if a consistent value *can* be returned it will be returned. For example, in the following code, The RGB code for pink (`#FFC0CB`) and `true` will be logged to the console because both the ranges in the `RangeAreas` object have a pink fill and both are entire columns.</span></span>

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

<span data-ttu-id="e3bb8-p129">As coisas se complicam quando a consistência não é possível. O comportamento das propriedades `RangeAreas` segue estes três princípios:</span><span class="sxs-lookup"><span data-stu-id="e3bb8-p129">Things get more complicated when consistency isn't possible. The behavior of `RangeAreas` properties follows these three principles:</span></span>

- <span data-ttu-id="e3bb8-258">Uma propriedade booleana de um objeto  `RangeAreas` retorna `false` , a menos que a propriedade seja verdadeira (true) para todos os intervalos membros.</span><span class="sxs-lookup"><span data-stu-id="e3bb8-258">A boolean property of a `RangeAreas` object returns `false` unless the property is true for all the member ranges.</span></span>
- <span data-ttu-id="e3bb8-259">Propriedades não-booleanas, com exceção da propriedade `address` , retornam `null` , a menos que a propriedade correspondente em todos os intervalos membro tenha o mesmo valor.</span><span class="sxs-lookup"><span data-stu-id="e3bb8-259">Non-boolean properties, with the exception of the `address` property, return `null` unless the corresponding property on all the member ranges has the same value.</span></span>
- <span data-ttu-id="e3bb8-260">A propriedade  `address` retornará uma sequência de caracteres delimitada por vírgulas dos endereços dos intervalos membros.</span><span class="sxs-lookup"><span data-stu-id="e3bb8-260">The `address` property returns a comma-delimited string of the addresses of the member ranges.</span></span>

<span data-ttu-id="e3bb8-p130">Por exemplo, o código a seguir cria um `RangeAreas` em que somente um intervalo é uma coluna inteira e apenas um é preenchido com rosa. O console mostrará `null` para a cor de preenchimento, `false` para a propriedade `isEntireRow` e "Sheet1!F3:F5, Sheet1!H:H"(supondo que o nome da planilha seja "Sheet1") para a propriedade `address`.</span><span class="sxs-lookup"><span data-stu-id="e3bb8-p130">For example, the following code creates a `RangeAreas` in which only one range is an entire column and only one is filled with pink. The console will show `null` for the fill color, `false` for the `isEntireRow` property, and "Sheet1!F3:F5, Sheet1!H:H" (assuming the sheet name is "Sheet1") for the `address` property.</span></span> 

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

## <a name="see-also"></a><span data-ttu-id="e3bb8-263">Confira também</span><span class="sxs-lookup"><span data-stu-id="e3bb8-263">See also</span></span>

- [<span data-ttu-id="e3bb8-264">Conceitos fundamentais de programação com a API JavaScript do Excel</span><span class="sxs-lookup"><span data-stu-id="e3bb8-264">Fundamental programming concepts with the Excel JavaScript API</span></span>](https://docs.microsoft.com/office/dev/add-ins/reference/overview/excel-add-ins-reference-overview)
- [<span data-ttu-id="e3bb8-265">Objeto Range (API JavaScript para Excel)</span><span class="sxs-lookup"><span data-stu-id="e3bb8-265">Range Object (JavaScript API for Excel)</span></span>](https://docs.microsoft.com/javascript/api/excel/excel.range)
- <span data-ttu-id="e3bb8-p131">[Objeto RangeAreas (API JavaScript para Excel)](https://docs.microsoft.com/javascript/api/excel/excel.rangeareas) (Este link pode não funcionar enquanto a API estiver na versão prévia. Como alternativa, consulte [office.d.ts beta](https://appsforoffice.microsoft.com/lib/beta/hosted/office.d.ts)).</span><span class="sxs-lookup"><span data-stu-id="e3bb8-p131">[RangeAreas Object (JavaScript API for Excel)](https://docs.microsoft.com/javascript/api/excel/excel.rangeareas) (This link may not work while the API is in preview. As an alternative, see [beta office.d.ts](https://appsforoffice.microsoft.com/lib/beta/hosted/office.d.ts).)</span></span>