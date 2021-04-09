---
title: Encontre células especiais em um intervalo usando a API JavaScript do Excel
description: Saiba como usar a API JavaScript do Excel para encontrar células especiais, como células com fórmulas, erros ou números.
ms.date: 04/02/2021
ms.prod: excel
localization_priority: Normal
ms.openlocfilehash: 6504873bcd8ab50bd4c03fe4f54b71d0bd920c5b
ms.sourcegitcommit: 54fef33bfc7d18a35b3159310bbd8b1c8312f845
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 04/09/2021
ms.locfileid: "51652758"
---
# <a name="find-special-cells-within-a-range-using-the-excel-javascript-api"></a><span data-ttu-id="01a8d-103">Encontre células especiais em um intervalo usando a API JavaScript do Excel</span><span class="sxs-lookup"><span data-stu-id="01a8d-103">Find special cells within a range using the Excel JavaScript API</span></span>

<span data-ttu-id="01a8d-104">Este artigo fornece exemplos de código que encontram células especiais em um intervalo usando a API JavaScript do Excel.</span><span class="sxs-lookup"><span data-stu-id="01a8d-104">This article provides code samples that find special cells within a range using the Excel JavaScript API.</span></span> <span data-ttu-id="01a8d-105">Para ver a lista completa de propriedades e métodos compatíveis com o `Range` objeto, consulte [Classe Excel.Range](/javascript/api/excel/excel.range).</span><span class="sxs-lookup"><span data-stu-id="01a8d-105">For the complete list of properties and methods that the `Range` object supports, see [Excel.Range class](/javascript/api/excel/excel.range).</span></span>

## <a name="find-ranges-with-special-cells"></a><span data-ttu-id="01a8d-106">Encontrar intervalos com células especiais</span><span class="sxs-lookup"><span data-stu-id="01a8d-106">Find ranges with special cells</span></span>

<span data-ttu-id="01a8d-107">Os [métodos Range.getSpecialCells](/javascript/api/excel/excel.range#getspecialcells-celltype--cellvaluetype-) e [Range.getSpecialCellsOrNullObject](/javascript/api/excel/excel.range#getspecialcellsornullobject-celltype--cellvaluetype-) encontram intervalos com base nas características de suas células e nos tipos de valores de suas células.</span><span class="sxs-lookup"><span data-stu-id="01a8d-107">The [Range.getSpecialCells](/javascript/api/excel/excel.range#getspecialcells-celltype--cellvaluetype-) and [Range.getSpecialCellsOrNullObject](/javascript/api/excel/excel.range#getspecialcellsornullobject-celltype--cellvaluetype-) methods find ranges based on the characteristics of their cells and the types of values of their cells.</span></span> <span data-ttu-id="01a8d-108">Os dois métodos retornam `RangeAreas` objetos.</span><span class="sxs-lookup"><span data-stu-id="01a8d-108">Both of these methods return `RangeAreas` objects.</span></span> <span data-ttu-id="01a8d-109">Aqui estão as assinaturas dos métodos do arquivo de tipos de dados TypeScript:</span><span class="sxs-lookup"><span data-stu-id="01a8d-109">Here are the signatures of the methods from the TypeScript data types file:</span></span>

```typescript
getSpecialCells(cellType: Excel.SpecialCellType, cellValueType?: Excel.SpecialCellValueType): Excel.RangeAreas;
```

```typescript
getSpecialCellsOrNullObject(cellType: Excel.SpecialCellType, cellValueType?: Excel.SpecialCellValueType): Excel.RangeAreas;
```

<span data-ttu-id="01a8d-110">O exemplo de código a seguir usa `getSpecialCells` o método para encontrar todas as células com fórmulas.</span><span class="sxs-lookup"><span data-stu-id="01a8d-110">The following code sample uses the `getSpecialCells` method to find all the cells with formulas.</span></span> <span data-ttu-id="01a8d-111">Sobre este código, observe:</span><span class="sxs-lookup"><span data-stu-id="01a8d-111">About this code, note:</span></span>

- <span data-ttu-id="01a8d-112">Ele limita a parte da planilha que precisa ser pesquisada chamando primeiro `Worksheet.getUsedRange` e chamando `getSpecialCells` para apenas esse intervalo.</span><span class="sxs-lookup"><span data-stu-id="01a8d-112">It limits the part of the sheet that needs to be searched by first calling `Worksheet.getUsedRange` and calling `getSpecialCells` for only that range.</span></span>
- <span data-ttu-id="01a8d-113">O método `getSpecialCells` retorna um objeto `RangeAreas`, então todas as células com fórmulas serão coloridas de rosa, mesmo que não sejam todas contíguas.</span><span class="sxs-lookup"><span data-stu-id="01a8d-113">The `getSpecialCells` method returns a `RangeAreas` object, so all of the cells with formulas will be colored pink even if they are not all contiguous.</span></span>

```js
Excel.run(function (context) {
    var sheet = context.workbook.worksheets.getActiveWorksheet();
    var usedRange = sheet.getUsedRange();
    var formulaRanges = usedRange.getSpecialCells(Excel.SpecialCellType.formulas);
    formulaRanges.format.fill.color = "pink";

    return context.sync();
})
```

<span data-ttu-id="01a8d-114">Se nenhuma célula com característica destino existe no intervalo, `getSpecialCells` exibe um erro **ItemNotFound**.</span><span class="sxs-lookup"><span data-stu-id="01a8d-114">If no cells with the targeted characteristic exist in the range, `getSpecialCells` throws an **ItemNotFound** error.</span></span> <span data-ttu-id="01a8d-115">Isso desvia o fluxo de controle para um `catch` bloco, se houver um.</span><span class="sxs-lookup"><span data-stu-id="01a8d-115">This diverts the flow of control to a `catch` block, if there is one.</span></span> <span data-ttu-id="01a8d-116">Se não houver um `catch` bloco, o erro interromperá o método.</span><span class="sxs-lookup"><span data-stu-id="01a8d-116">If there isn't a `catch` block, the error halts the method.</span></span>

<span data-ttu-id="01a8d-117">Se você espera que células com característica direcionada sempre deveriam existir, provavelmente desejará o código para gerar um erro se as células não estiverem lá.</span><span class="sxs-lookup"><span data-stu-id="01a8d-117">If you expect that cells with the targeted characteristic should always exist, you'll likely want your code to throw an error if those cells aren't there.</span></span> <span data-ttu-id="01a8d-118">Se for um cenário válido que não há uma ou mais células correspondentes, o código deve verificar se há essa possibilidade e tratar normalmente sem enviar um erro.</span><span class="sxs-lookup"><span data-stu-id="01a8d-118">If it's a valid scenario that there aren't any matching cells, your code should check for this possibility and handle it gracefully without throwing an error.</span></span> <span data-ttu-id="01a8d-119">Você pode obter esse comportamento com o `getSpecialCellsOrNullObject` método e sua propriedade retornada `isNullObject`.</span><span class="sxs-lookup"><span data-stu-id="01a8d-119">You can achieve this behavior with the `getSpecialCellsOrNullObject` method and its returned `isNullObject` property.</span></span> <span data-ttu-id="01a8d-120">O exemplo de código a seguir usa esse padrão.</span><span class="sxs-lookup"><span data-stu-id="01a8d-120">The following code sample uses this pattern.</span></span> <span data-ttu-id="01a8d-121">Sobre este código, observe:</span><span class="sxs-lookup"><span data-stu-id="01a8d-121">About this code, note:</span></span>

- <span data-ttu-id="01a8d-122">O `getSpecialCellsOrNullObject` método sempre retorna um objeto proxy, portanto, nunca está no sentido `null` javaScript comum.</span><span class="sxs-lookup"><span data-stu-id="01a8d-122">The `getSpecialCellsOrNullObject` method always returns a proxy object, so it's never `null` in the ordinary JavaScript sense.</span></span> <span data-ttu-id="01a8d-123">Mas se nenhuma célula de correspondência for encontrada, as propriedades do objeto`isNullObject` serão definida como `true`.</span><span class="sxs-lookup"><span data-stu-id="01a8d-123">But if no matching cells are found, the `isNullObject` property of the object is set to `true`.</span></span>
- <span data-ttu-id="01a8d-124">Ele chama `context.sync` *antes* de testar a propriedade`isNullObject`.</span><span class="sxs-lookup"><span data-stu-id="01a8d-124">It calls `context.sync` *before* it tests the `isNullObject` property.</span></span> <span data-ttu-id="01a8d-125">Esse é um requisito com todos os métodos e propriedades `*OrNullObject`, pois sempre terá que carregar e sincronizar as propriedades na ordem para lê-la.</span><span class="sxs-lookup"><span data-stu-id="01a8d-125">This is a requirement with all `*OrNullObject` methods and properties, because you always have to load and sync a property in order to read it.</span></span> <span data-ttu-id="01a8d-126">No entanto, não é necessário carregar *explicitamente* a `isNullObject` propriedade.</span><span class="sxs-lookup"><span data-stu-id="01a8d-126">However, it's not necessary to *explicitly* load the `isNullObject` property.</span></span> <span data-ttu-id="01a8d-127">Ele é carregado automaticamente pelo `context.sync` mesmo se não for chamado no `load` objeto.</span><span class="sxs-lookup"><span data-stu-id="01a8d-127">It's automatically loaded by the `context.sync` even if `load` is not called on the object.</span></span> <span data-ttu-id="01a8d-128">Para obter mais informações, consulte Métodos e propriedades [ \* OrNullObject](../develop/application-specific-api-model.md#ornullobject-methods-and-properties).</span><span class="sxs-lookup"><span data-stu-id="01a8d-128">For more information, see [\*OrNullObject methods and properties](../develop/application-specific-api-model.md#ornullobject-methods-and-properties).</span></span>
- <span data-ttu-id="01a8d-129">Você pode testar esse código selecionando primeiro um intervalo sem células de fórmula e executando-o.</span><span class="sxs-lookup"><span data-stu-id="01a8d-129">You can test this code by first selecting a range that has no formula cells and running it.</span></span> <span data-ttu-id="01a8d-130">Selecione um intervalo que tem pelo menos uma célula com uma fórmula e execute novamente.</span><span class="sxs-lookup"><span data-stu-id="01a8d-130">Then select a range that has at least one cell with a formula and run it again.</span></span>

```js
Excel.run(function (context) {
    var range = context.workbook.getSelectedRange();
    var formulaRanges = range.getSpecialCellsOrNullObject(Excel.SpecialCellType.formulas);
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

<span data-ttu-id="01a8d-131">Para simplificar, todos os outros exemplos de código neste artigo usam o `getSpecialCells` método em vez de  `getSpecialCellsOrNullObject` .</span><span class="sxs-lookup"><span data-stu-id="01a8d-131">For simplicity, all other code samples in this article use the `getSpecialCells` method instead of  `getSpecialCellsOrNullObject`.</span></span>

## <a name="narrow-the-target-cells-with-cell-value-types"></a><span data-ttu-id="01a8d-132">Restrinja as células de destino com tipos de valor de célula</span><span class="sxs-lookup"><span data-stu-id="01a8d-132">Narrow the target cells with cell value types</span></span>

<span data-ttu-id="01a8d-133">As `Range.getSpecialCells()` e `Range.getSpecialCellsOrNullObject()` métodos aceitam um segundo parâmetro opcional usado para restringir ainda mais as células de destino.</span><span class="sxs-lookup"><span data-stu-id="01a8d-133">The `Range.getSpecialCells()` and `Range.getSpecialCellsOrNullObject()` methods accept an optional second parameter used to further narrow down the targeted cells.</span></span> <span data-ttu-id="01a8d-134">Este segundo parâmetro é uma `Excel.SpecialCellValueType` você usar para especificar que você quer apenas células que contêm determinados tipos de valores.</span><span class="sxs-lookup"><span data-stu-id="01a8d-134">This second parameter is an `Excel.SpecialCellValueType` you use to specify that you only want cells that contain certain types of values.</span></span>

> [!NOTE]
> <span data-ttu-id="01a8d-135">O `Excel.SpecialCellValueType` parâmetro só pode ser usado se a `Excel.SpecialCellType` está `Excel.SpecialCellType.formulas` ou `Excel.SpecialCellType.constants`.</span><span class="sxs-lookup"><span data-stu-id="01a8d-135">The `Excel.SpecialCellValueType` parameter can only be used if the `Excel.SpecialCellType` is `Excel.SpecialCellType.formulas` or `Excel.SpecialCellType.constants`.</span></span>

### <a name="test-for-a-single-cell-value-type"></a><span data-ttu-id="01a8d-136">Teste para um tipo de valor da célula única</span><span class="sxs-lookup"><span data-stu-id="01a8d-136">Test for a single cell value type</span></span>

<span data-ttu-id="01a8d-137">O `Excel.SpecialCellValueType` enumeração com esses quatro tipos básicos (além dos outros valores combinados descritos nesta seção posterior):</span><span class="sxs-lookup"><span data-stu-id="01a8d-137">The `Excel.SpecialCellValueType` enum has these four basic types (in addition to the other combined values described later in this section):</span></span>

- `Excel.SpecialCellValueType.errors`
- <span data-ttu-id="01a8d-138">`Excel.SpecialCellValueType.logical` (ou seja, booliano)</span><span class="sxs-lookup"><span data-stu-id="01a8d-138">`Excel.SpecialCellValueType.logical` (which means boolean)</span></span>
- `Excel.SpecialCellValueType.numbers`
- `Excel.SpecialCellValueType.text`

<span data-ttu-id="01a8d-139">O exemplo de código a seguir localiza células especiais que são constantes numéricas e colore essas células rosa.</span><span class="sxs-lookup"><span data-stu-id="01a8d-139">The following code sample finds special cells that are numerical constants and colors those cells pink.</span></span> <span data-ttu-id="01a8d-140">Sobre este código, observe:</span><span class="sxs-lookup"><span data-stu-id="01a8d-140">About this code, note:</span></span>

- <span data-ttu-id="01a8d-141">Ele só realça células que têm um valor de número literal.</span><span class="sxs-lookup"><span data-stu-id="01a8d-141">It only highlights cells that have a literal number value.</span></span> <span data-ttu-id="01a8d-142">Ele não realça células que têm uma fórmula (mesmo que o resultado seja um número) ou um booleano, texto ou células de estado de erro.</span><span class="sxs-lookup"><span data-stu-id="01a8d-142">It won't highlight cells that have a formula (even if the result is a number) or a boolean, text, or error state cells.</span></span>
- <span data-ttu-id="01a8d-143">Para testar o código, certifique-se de que a planilha tenha algumas células com valores numéricos literais, algumas com outros tipos de valores literais e algumas com fórmulas.</span><span class="sxs-lookup"><span data-stu-id="01a8d-143">To test the code, be sure the worksheet has some cells with literal number values, some with other kinds of literal values, and some with formulas.</span></span>

```js
Excel.run(function (context) {
    var sheet = context.workbook.worksheets.getActiveWorksheet();
    var usedRange = sheet.getUsedRange();
    var constantNumberRanges = usedRange.getSpecialCells(
        Excel.SpecialCellType.constants,
        Excel.SpecialCellValueType.numbers);
    constantNumberRanges.format.fill.color = "pink";

    return context.sync();
})
```

### <a name="test-for-multiple-cell-value-types"></a><span data-ttu-id="01a8d-144">Teste para vários tipos de valores de célula</span><span class="sxs-lookup"><span data-stu-id="01a8d-144">Test for multiple cell value types</span></span>

<span data-ttu-id="01a8d-145">Às vezes, você precisa operar com mais de um tipo de valor de célula, como todas as células com valor de texto e com valor booliano ("Lógico"). (`Excel.SpecialCellValueType.logical`).</span><span class="sxs-lookup"><span data-stu-id="01a8d-145">Sometimes you need to operate on more than one cell value type, such as all text-valued and all boolean-valued (`Excel.SpecialCellValueType.logical`) cells.</span></span> <span data-ttu-id="01a8d-146">O `Excel.SpecialCellValueType` enumeração tem valores com tipos combinado.</span><span class="sxs-lookup"><span data-stu-id="01a8d-146">The `Excel.SpecialCellValueType` enum has values with combined types.</span></span> <span data-ttu-id="01a8d-147">Por exemplo, `Excel.SpecialCellValueType.logicalText` segmentará todas as células boolianas e todos os valores de texto.</span><span class="sxs-lookup"><span data-stu-id="01a8d-147">For example, `Excel.SpecialCellValueType.logicalText` targets all boolean and all text-valued cells.</span></span> <span data-ttu-id="01a8d-148">`Excel.SpecialCellValueType.all` é o valor padrão, que não limita os tipos de valor da célula retornados.</span><span class="sxs-lookup"><span data-stu-id="01a8d-148">`Excel.SpecialCellValueType.all` is the default value, which does not limit the cell value types returned.</span></span> <span data-ttu-id="01a8d-149">O exemplo de código a seguir colore todas as células com fórmulas que produzem número ou valor booleano.</span><span class="sxs-lookup"><span data-stu-id="01a8d-149">The following code sample colors all cells with formulas that produce number or boolean value.</span></span>

```js
Excel.run(function (context) {
    var sheet = context.workbook.worksheets.getActiveWorksheet();
    var usedRange = sheet.getUsedRange();
    var formulaLogicalNumberRanges = usedRange.getSpecialCells(
        Excel.SpecialCellType.formulas,
        Excel.SpecialCellValueType.logicalNumbers);
    formulaLogicalNumberRanges.format.fill.color = "pink";

    return context.sync();
})
```

## <a name="see-also"></a><span data-ttu-id="01a8d-150">Confira também</span><span class="sxs-lookup"><span data-stu-id="01a8d-150">See also</span></span>

- [<span data-ttu-id="01a8d-151">Modelo de objeto JavaScript do Excel em Suplementos do Office</span><span class="sxs-lookup"><span data-stu-id="01a8d-151">Excel JavaScript object model in Office Add-ins</span></span>](excel-add-ins-core-concepts.md)
- [<span data-ttu-id="01a8d-152">Trabalhar com células usando a API JavaScript do Excel</span><span class="sxs-lookup"><span data-stu-id="01a8d-152">Work with cells using the Excel JavaScript API</span></span>](excel-add-ins-cells.md)
- [<span data-ttu-id="01a8d-153">Encontre uma cadeia de caracteres usando a API JavaScript do Excel</span><span class="sxs-lookup"><span data-stu-id="01a8d-153">Find a string using the Excel JavaScript API</span></span>](excel-add-ins-ranges-string-match.md)
- [<span data-ttu-id="01a8d-154">Trabalhar simultaneamente com vários intervalos em suplementos do Excel</span><span class="sxs-lookup"><span data-stu-id="01a8d-154">Work with multiple ranges simultaneously in Excel add-ins</span></span>](excel-add-ins-multiple-ranges.md)
