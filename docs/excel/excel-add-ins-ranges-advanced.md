---
title: Trabalhar com intervalos usando a API JavaScript do Excel (avançado)
description: ''
ms.date: 04/15/2019
localization_priority: Normal
ms.openlocfilehash: aacbe930e2cf3da4d10b61bfe8f34efe1094c113
ms.sourcegitcommit: 6d375518c119d09c8d3fb5f0cc4583ba5b20ac03
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 04/18/2019
ms.locfileid: "31914235"
---
# <a name="work-with-ranges-using-the-excel-javascript-api-advanced"></a><span data-ttu-id="0cf9d-102">Trabalhar com intervalos usando a API JavaScript do Excel (avançado)</span><span class="sxs-lookup"><span data-stu-id="0cf9d-102">Work with ranges using the Excel JavaScript API (advanced)</span></span>

<span data-ttu-id="0cf9d-103">Este artigo baseia-se em informações em [Trabalhar com intervalos usando a API JavaScript do Excel (fundamental)](excel-add-ins-ranges.md) fornecendo exemplos de código que mostram como executar tarefas mais avançadas com intervalos usando a API JavaScript do Excel.</span><span class="sxs-lookup"><span data-stu-id="0cf9d-103">This article builds upon information in [Work with ranges using the Excel JavaScript API (fundamental)](excel-add-ins-ranges.md) by providing code samples that show how to perform more advanced tasks with ranges using the Excel JavaScript API.</span></span> <span data-ttu-id="0cf9d-104">Para obter a lista completa de propriedades e métodos que o objeto **Range** suporta, confira [Objeto Range (API JavaScript para Excel)](/javascript/api/excel/excel.range).</span><span class="sxs-lookup"><span data-stu-id="0cf9d-104">For the complete list of properties and methods that the **Range** object supports, see [Range Object (JavaScript API for Excel)](/javascript/api/excel/excel.range).</span></span>

## <a name="work-with-dates-using-the-moment-msdate-plug-in"></a><span data-ttu-id="0cf9d-105">Trabalhar com datas usando o plug-in Moment-MSDate</span><span class="sxs-lookup"><span data-stu-id="0cf9d-105">Work with dates using the Moment-MSDate plug-in</span></span>

<span data-ttu-id="0cf9d-106">A [biblioteca Moment do JavaScript](https://momentjs.com/) fornece uma maneira conveniente de usar datas e carimbos de data e hora.</span><span class="sxs-lookup"><span data-stu-id="0cf9d-106">The [Moment JavaScript library](https://momentjs.com/) provides a convenient way to use dates and timestamps.</span></span> <span data-ttu-id="0cf9d-107">O [plug-in Moment-MSDate](https://www.npmjs.com/package/moment-msdate) converte o formato de momentos em um formato mais apropriado para o Excel.</span><span class="sxs-lookup"><span data-stu-id="0cf9d-107">The [Moment-MSDate plug-in](https://www.npmjs.com/package/moment-msdate) converts the format of moments into one preferable for Excel.</span></span> <span data-ttu-id="0cf9d-108">Este é o mesmo formato que a [função NOW](https://support.office.com/article/now-function-3337fd29-145a-4347-b2e6-20c904739c46) retorna.</span><span class="sxs-lookup"><span data-stu-id="0cf9d-108">This is the same format the [NOW function](https://support.office.com/article/now-function-3337fd29-145a-4347-b2e6-20c904739c46) returns.</span></span>

<span data-ttu-id="0cf9d-109">O código a seguir mostra como definir o intervalo em \*\* B4 \*\* para o carimbo de data/hora de um momento:</span><span class="sxs-lookup"><span data-stu-id="0cf9d-109">The following code shows how to set the range at **B4** to a moment's timestamp:</span></span>

```js
Excel.run(function (context) {
    var sheet = context.workbook.worksheets.getItem("Sample");

    var now = Date.now();
    var nowMoment = moment(now);
    var nowMS = nowMoment.toOADate();

    var dateRange = sheet.getRange("B4");
    dateRange.values = [[nowMS]];

    dateRange.numberFormat = [["[$-409]m/d/yy h:mm AM/PM;@"]];

    return context.sync();
}).catch(errorHandlerFunction);
```

<span data-ttu-id="0cf9d-110">É uma técnica semelhante para retirar a data da célula e convertê-la em um momento ou outro formato, conforme demonstrado no código a seguir:</span><span class="sxs-lookup"><span data-stu-id="0cf9d-110">It is a similar technique to get the date back out of the cell and convert it to a moment or other format, as demonstrated in the following code:</span></span>

```js
Excel.run(function (context) {
    var sheet = context.workbook.worksheets.getItem("Sample");

    var dateRange = sheet.getRange("B4");
    dateRange.load("values");

    return context.sync().then(function () {
        var nowMS = dateRange.values[0][0];

        // log the date as a moment
        var nowMoment = moment.fromOADate(nowMS);
        console.log(`get (moment): ${JSON.stringify(nowMoment)}`);

        // log the date as a UNIX-style timestamp
        var now = nowMoment.unix();
        console.log(`get (timestamp): ${now}`);
    });
}).catch(errorHandlerFunction);
```

<span data-ttu-id="0cf9d-111">Seu suplemento terá que formatar os intervalos para exibir as datas em um formato mais legível.</span><span class="sxs-lookup"><span data-stu-id="0cf9d-111">Your add-in will have to format the ranges to display the dates in a more human-readable form.</span></span> <span data-ttu-id="0cf9d-112">O exemplo de `"[$-409]m/d/yy h:mm AM/PM;@"` exibe a hora como "3/12/18 15:57".</span><span class="sxs-lookup"><span data-stu-id="0cf9d-112">The example of `"[$-409]m/d/yy h:mm AM/PM;@"` displays a time like "12/3/18 3:57 PM".</span></span> <span data-ttu-id="0cf9d-113">Para obter mais informações sobre formatos de números de data e hora, confira as "Diretrizes para formatos de data e hora" no artigo [Diretrizes de revisão para personalizar um formato de número](https://support.office.com/article/review-guidelines-for-customizing-a-number-format-c0a1d1fa-d3f4-4018-96b7-9c9354dd99f5).</span><span class="sxs-lookup"><span data-stu-id="0cf9d-113">For more information about date and time number formats, please see the "Guidelines for date and time formats" in the [Review guidelines for customizing a number format](https://support.office.com/article/review-guidelines-for-customizing-a-number-format-c0a1d1fa-d3f4-4018-96b7-9c9354dd99f5) article.</span></span>

## <a name="work-with-multiple-ranges-simultaneously-preview"></a><span data-ttu-id="0cf9d-114">Trabalhar simultaneamente com vários intervalos (Visualização)</span><span class="sxs-lookup"><span data-stu-id="0cf9d-114">Work with multiple ranges simultaneously (preview)</span></span>

> [!NOTE]
> <span data-ttu-id="0cf9d-115">O `RangeAreas` objeto está disponível atualmente apenas na visualização pública.</span><span class="sxs-lookup"><span data-stu-id="0cf9d-115">The `RangeAreas` object is currently available only in public preview.</span></span> [!INCLUDE [Information about using preview APIs](../includes/using-excel-preview-apis.md)]

<span data-ttu-id="0cf9d-116">O `RangeAreas` objeto permite ao suplemento executar operações em vários intervalos de uma só vez.</span><span class="sxs-lookup"><span data-stu-id="0cf9d-116">The `RangeAreas` object lets your add-in perform operations on multiple ranges at once.</span></span> <span data-ttu-id="0cf9d-117">Esses intervalos poderão ser contíguos, mas não precisam ser.</span><span class="sxs-lookup"><span data-stu-id="0cf9d-117">These ranges may be contiguous, but do not have to be.</span></span> <span data-ttu-id="0cf9d-118">`RangeAreas` são descritas ainda mais no artigo [Trabalhar com vários intervalos simultaneamente em suplementos do Excel](excel-add-ins-multiple-ranges.md).</span><span class="sxs-lookup"><span data-stu-id="0cf9d-118">`RangeAreas` are further discussed in the article [Work with multiple ranges simultaneously in Excel add-ins](excel-add-ins-multiple-ranges.md).</span></span>

## <a name="find-special-cells-within-a-range-preview"></a><span data-ttu-id="0cf9d-119">Localizar células especiais em um intervalo (visualização)</span><span class="sxs-lookup"><span data-stu-id="0cf9d-119">Find special cells within a range (preview)</span></span>

> [!NOTE]
> <span data-ttu-id="0cf9d-120">Os `getSpecialCells` métodos `getSpecialCellsOrNullObject` e estão atualmente disponíveis somente na visualização pública.</span><span class="sxs-lookup"><span data-stu-id="0cf9d-120">The `getSpecialCells` and `getSpecialCellsOrNullObject` methods are currently available only in public preview.</span></span> [!INCLUDE [Information about using preview APIs](../includes/using-excel-preview-apis.md)]

<span data-ttu-id="0cf9d-121">As `Range.getSpecialCells()` e `Range.getSpecialCellsOrNullObject()` métodos localizar intervalos com base nas características de suas células e os tipos de valores de suas células.</span><span class="sxs-lookup"><span data-stu-id="0cf9d-121">The `Range.getSpecialCells()` and `Range.getSpecialCellsOrNullObject()` methods find ranges based on the characteristics of their cells and the types of values of their cells.</span></span> <span data-ttu-id="0cf9d-122">Os dois métodos retornam `RangeAreas` objetos.</span><span class="sxs-lookup"><span data-stu-id="0cf9d-122">Both of these methods return `RangeAreas` objects.</span></span> <span data-ttu-id="0cf9d-123">Aqui estão as assinaturas dos métodos do arquivo de tipos de dados TypeScript:</span><span class="sxs-lookup"><span data-stu-id="0cf9d-123">Here are the signatures of the methods from the TypeScript data types file:</span></span>

```typescript
getSpecialCells(cellType: Excel.SpecialCellType, cellValueType?: Excel.SpecialCellValueType): Excel.RangeAreas;
```

```typescript
getSpecialCellsOrNullObject(cellType: Excel.SpecialCellType, cellValueType?: Excel.SpecialCellValueType): Excel.RangeAreas;
```

<span data-ttu-id="0cf9d-124">O exemplo a seguir usa o `getSpecialCells` método para localizar células com fórmulas.</span><span class="sxs-lookup"><span data-stu-id="0cf9d-124">The following example uses the `getSpecialCells` method to find all the cells with formulas.</span></span> <span data-ttu-id="0cf9d-125">Sobre este código, observe:</span><span class="sxs-lookup"><span data-stu-id="0cf9d-125">About this code, note:</span></span>

- <span data-ttu-id="0cf9d-126">Ele limita a parte da planilha que precisa ser pesquisada chamando primeiro `Worksheet.getUsedRange` e chamando `getSpecialCells` para apenas esse intervalo.</span><span class="sxs-lookup"><span data-stu-id="0cf9d-126">It limits the part of the sheet that needs to be searched by first calling `Worksheet.getUsedRange` and calling `getSpecialCells` for only that range.</span></span>
- <span data-ttu-id="0cf9d-127">O método `getSpecialCells` retorna um objeto `RangeAreas`, então todas as células com fórmulas serão coloridas de rosa, mesmo que não sejam todas contíguas.</span><span class="sxs-lookup"><span data-stu-id="0cf9d-127">The `getSpecialCells` method returns a `RangeAreas` object, so all of the cells with formulas will be colored pink even if they are not all contiguous.</span></span>

```js
Excel.run(function (context) {
    var sheet = context.workbook.worksheets.getActiveWorksheet();
    var usedRange = sheet.getUsedRange();
    var formulaRanges = usedRange.getSpecialCells(Excel.SpecialCellType.formulas);
    formulaRanges.format.fill.color = "pink";

    return context.sync();
})
```

<span data-ttu-id="0cf9d-128">Se nenhuma célula com característica destino existe no intervalo, `getSpecialCells` exibe um erro **ItemNotFound**.</span><span class="sxs-lookup"><span data-stu-id="0cf9d-128">If no cells with the targeted characteristic exist in the range, `getSpecialCells` throws an **ItemNotFound** error.</span></span> <span data-ttu-id="0cf9d-129">Isso desvia o fluxo de controle para um `catch` bloco, se houver um.</span><span class="sxs-lookup"><span data-stu-id="0cf9d-129">This diverts the flow of control to a `catch` block, if there is one.</span></span> <span data-ttu-id="0cf9d-130">Se não houver um `catch` bloco, o erro interrompe a função.</span><span class="sxs-lookup"><span data-stu-id="0cf9d-130">If there isn't a `catch` block, the error halts the function.</span></span>

<span data-ttu-id="0cf9d-131">Se você espera que células com característica direcionada sempre deveriam existir, provavelmente desejará o código para gerar um erro se as células não estiverem lá.</span><span class="sxs-lookup"><span data-stu-id="0cf9d-131">If you expect that cells with the targeted characteristic should always exist, you'll likely want your code to throw an error if those cells aren't there.</span></span> <span data-ttu-id="0cf9d-132">Se for um cenário válido que não há uma ou mais células correspondentes, o código deve verificar se há essa possibilidade e tratar normalmente sem enviar um erro.</span><span class="sxs-lookup"><span data-stu-id="0cf9d-132">If it's a valid scenario that there aren't any matching cells, your code should check for this possibility and handle it gracefully without throwing an error.</span></span> <span data-ttu-id="0cf9d-133">Você pode obter esse comportamento com o `getSpecialCellsOrNullObject` método e sua propriedade retornada `isNullObject`.</span><span class="sxs-lookup"><span data-stu-id="0cf9d-133">You can achieve this behavior with the `getSpecialCellsOrNullObject` method and its returned `isNullObject` property.</span></span> <span data-ttu-id="0cf9d-134">O exemplo a seguir usa esse padrão.</span><span class="sxs-lookup"><span data-stu-id="0cf9d-134">The following example uses this pattern.</span></span> <span data-ttu-id="0cf9d-135">Sobre este código, observe:</span><span class="sxs-lookup"><span data-stu-id="0cf9d-135">About this code, note:</span></span>

- <span data-ttu-id="0cf9d-136">O método `getSpecialCellsOrNullObject` sempre retorna um objeto de proxy, portanto, `null` nunca está no sentido comum do JavaScript.</span><span class="sxs-lookup"><span data-stu-id="0cf9d-136">The `getSpecialCellsOrNullObject` method always returns a proxy object, so it is never `null` in the ordinary JavaScript sense.</span></span> <span data-ttu-id="0cf9d-137">Mas se nenhuma célula de correspondência for encontrada, as propriedades do objeto`isNullObject` serão definida como `true`.</span><span class="sxs-lookup"><span data-stu-id="0cf9d-137">But if no matching cells are found, the `isNullObject` property of the object is set to `true`.</span></span>
- <span data-ttu-id="0cf9d-138">Ele chama `context.sync` *antes* de testar a propriedade`isNullObject`.</span><span class="sxs-lookup"><span data-stu-id="0cf9d-138">It calls `context.sync` *before* it tests the `isNullObject` property.</span></span> <span data-ttu-id="0cf9d-139">Esse é um requisito com todos os métodos e propriedades `*OrNullObject`, pois sempre terá que carregar e sincronizar as propriedades na ordem para lê-la.</span><span class="sxs-lookup"><span data-stu-id="0cf9d-139">This is a requirement with all `*OrNullObject` methods and properties, because you always have to load and sync a property in order to read it.</span></span> <span data-ttu-id="0cf9d-140">No entanto, não é necessário carregar *explicitamente* a propriedade`isNullObject`.</span><span class="sxs-lookup"><span data-stu-id="0cf9d-140">However, it is not necessary to *explicitly* load the `isNullObject` property.</span></span> <span data-ttu-id="0cf9d-141">Será carregado automaticamente pelo `context.sync` mesmo se `load` não for chamado no objeto.</span><span class="sxs-lookup"><span data-stu-id="0cf9d-141">It is automatically loaded by the `context.sync` even if `load` is not called on the object.</span></span> <span data-ttu-id="0cf9d-142">Para saber mais, confira [ \*OrNullObject](/office/dev/add-ins/excel/excel-add-ins-advanced-concepts#ornullobject-methods).</span><span class="sxs-lookup"><span data-stu-id="0cf9d-142">For more information, see [\*OrNullObject](/office/dev/add-ins/excel/excel-add-ins-advanced-concepts#ornullobject-methods).</span></span>
- <span data-ttu-id="0cf9d-143">Você pode testar esse código selecionando primeiro um intervalo sem células de fórmula e executando-o.</span><span class="sxs-lookup"><span data-stu-id="0cf9d-143">You can test this code by first selecting a range that has no formula cells and running it.</span></span> <span data-ttu-id="0cf9d-144">Selecione um intervalo que tem pelo menos uma célula com uma fórmula e execute novamente.</span><span class="sxs-lookup"><span data-stu-id="0cf9d-144">Then select a range that has at least one cell with a formula and run it again.</span></span>

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

<span data-ttu-id="0cf9d-145">Para manter a simplicidade, todos os outros exemplos deste artigo usam o método `getSpecialCells` em vez de `getSpecialCellsOrNullObject`.</span><span class="sxs-lookup"><span data-stu-id="0cf9d-145">For simplicity, all other examples in this article use the `getSpecialCells` method instead of  `getSpecialCellsOrNullObject`.</span></span>

### <a name="narrow-the-target-cells-with-cell-value-types"></a><span data-ttu-id="0cf9d-146">Restrinja as células de destino com tipos de valor de célula</span><span class="sxs-lookup"><span data-stu-id="0cf9d-146">Narrow the target cells with cell value types</span></span>

<span data-ttu-id="0cf9d-147">As `Range.getSpecialCells()` e `Range.getSpecialCellsOrNullObject()` métodos aceitam um segundo parâmetro opcional usado para restringir ainda mais as células de destino.</span><span class="sxs-lookup"><span data-stu-id="0cf9d-147">The `Range.getSpecialCells()` and `Range.getSpecialCellsOrNullObject()` methods accept an optional second parameter used to further narrow down the targeted cells.</span></span> <span data-ttu-id="0cf9d-148">Este segundo parâmetro é uma `Excel.SpecialCellValueType` você usar para especificar que você quer apenas células que contêm determinados tipos de valores.</span><span class="sxs-lookup"><span data-stu-id="0cf9d-148">This second parameter is an `Excel.SpecialCellValueType` you use to specify that you only want cells that contain certain types of values.</span></span>

> [!NOTE]
> <span data-ttu-id="0cf9d-149">O `Excel.SpecialCellValueType` parâmetro só pode ser usado se a `Excel.SpecialCellType` está `Excel.SpecialCellType.formulas` ou `Excel.SpecialCellType.constants`.</span><span class="sxs-lookup"><span data-stu-id="0cf9d-149">The `Excel.SpecialCellValueType` parameter can only be used if the `Excel.SpecialCellType` is `Excel.SpecialCellType.formulas` or `Excel.SpecialCellType.constants`.</span></span>

#### <a name="test-for-a-single-cell-value-type"></a><span data-ttu-id="0cf9d-150">Teste para um tipo de valor da célula única</span><span class="sxs-lookup"><span data-stu-id="0cf9d-150">Test for a single cell value type</span></span>

<span data-ttu-id="0cf9d-151">O `Excel.SpecialCellValueType` enumeração com esses quatro tipos básicos (além dos outros valores combinados descritos nesta seção posterior):</span><span class="sxs-lookup"><span data-stu-id="0cf9d-151">The `Excel.SpecialCellValueType` enum has these four basic types (in addition to the other combined values described later in this section):</span></span>

- `Excel.SpecialCellValueType.errors`
- <span data-ttu-id="0cf9d-152">`Excel.SpecialCellValueType.logical` (ou seja, booliano)</span><span class="sxs-lookup"><span data-stu-id="0cf9d-152">`Excel.SpecialCellValueType.logical` (which means boolean)</span></span>
- `Excel.SpecialCellValueType.numbers`
- `Excel.SpecialCellValueType.text`

<span data-ttu-id="0cf9d-153">O exemplo a seguir localiza as células especiais que são constantes numéricos e colore essas células em rosa.</span><span class="sxs-lookup"><span data-stu-id="0cf9d-153">The following example finds special cells that are numerical constants and colors those cells pink.</span></span> <span data-ttu-id="0cf9d-154">Sobre este código, observe:</span><span class="sxs-lookup"><span data-stu-id="0cf9d-154">About this code, note:</span></span>

- <span data-ttu-id="0cf9d-155">Ele apenas irá realçar células que contêm um valor numérico literal.</span><span class="sxs-lookup"><span data-stu-id="0cf9d-155">It will only highlight cells that have a literal number value.</span></span> <span data-ttu-id="0cf9d-156">Ele não destacará as células que têm uma fórmula (mesmo se o resultado for um número) ou células de estado booliano, de texto ou de erro.</span><span class="sxs-lookup"><span data-stu-id="0cf9d-156">It will not highlight cells that have a formula (even if the result is a number) or a boolean, text, or error state cells.</span></span>
- <span data-ttu-id="0cf9d-157">Para testar o código, certifique-se de que a planilha tenha algumas células com valores numéricos literais, algumas com outros tipos de valores literais e algumas com fórmulas.</span><span class="sxs-lookup"><span data-stu-id="0cf9d-157">To test the code, be sure the worksheet has some cells with literal number values, some with other kinds of literal values, and some with formulas.</span></span>

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

#### <a name="test-for-multiple-cell-value-types"></a><span data-ttu-id="0cf9d-158">Teste para vários tipos de valores de célula</span><span class="sxs-lookup"><span data-stu-id="0cf9d-158">Test for multiple cell value types</span></span>

<span data-ttu-id="0cf9d-159">Às vezes, você precisa operar com mais de um tipo de valor de célula, como todas as células com valor de texto e com valor booliano ("Lógico"). (`Excel.SpecialCellValueType.logical`).</span><span class="sxs-lookup"><span data-stu-id="0cf9d-159">Sometimes you need to operate on more than one cell value type, such as all text-valued and all boolean-valued (`Excel.SpecialCellValueType.logical`) cells.</span></span> <span data-ttu-id="0cf9d-160">O `Excel.SpecialCellValueType` enumeração tem valores com tipos combinado.</span><span class="sxs-lookup"><span data-stu-id="0cf9d-160">The `Excel.SpecialCellValueType` enum has values with combined types.</span></span> <span data-ttu-id="0cf9d-161">Por exemplo, `Excel.SpecialCellValueType.logicalText` segmentará todas as células boolianas e todos os valores de texto.</span><span class="sxs-lookup"><span data-stu-id="0cf9d-161">For example, `Excel.SpecialCellValueType.logicalText` targets all boolean and all text-valued cells.</span></span> <span data-ttu-id="0cf9d-162">`Excel.SpecialCellValueType.all` é o valor padrão, que não limita os tipos de valor da célula retornados.</span><span class="sxs-lookup"><span data-stu-id="0cf9d-162">`Excel.SpecialCellValueType.all` is the default value, which does not limit the cell value types returned.</span></span> <span data-ttu-id="0cf9d-163">O exemplo a seguir destaca todas as células com fórmulas que produzem valores ou números boolianos.</span><span class="sxs-lookup"><span data-stu-id="0cf9d-163">The following example colors all cells with formulas that produce number or boolean value.</span></span>

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

## <a name="copy-and-paste-preview"></a><span data-ttu-id="0cf9d-164">Copiar e colar (visualização)</span><span class="sxs-lookup"><span data-stu-id="0cf9d-164">Copy and paste (preview)</span></span>

> [!NOTE]
> <span data-ttu-id="0cf9d-165">A função `Range.copyFrom` só está disponível atualmente na versão prévia pública.</span><span class="sxs-lookup"><span data-stu-id="0cf9d-165">The `Range.copyFrom` function is currently available only in public preview.</span></span> [!INCLUDE [Information about using preview APIs](../includes/using-excel-preview-apis.md)]

<span data-ttu-id="0cf9d-166">A função de `copyFrom` do intervalo replica o comportamento de copiar e colar da IU do Excel.</span><span class="sxs-lookup"><span data-stu-id="0cf9d-166">Range’s `copyFrom` function replicates the copy-and-paste behavior of the Excel UI.</span></span> <span data-ttu-id="0cf9d-167">O objeto de intervalo para o qual a função`copyFrom` é chamada é o destino.</span><span class="sxs-lookup"><span data-stu-id="0cf9d-167">The range object that `copyFrom` is called on is the destination.</span></span>
<span data-ttu-id="0cf9d-168">A fonte a ser copiada é passada como um intervalo ou um endereço de cadeia de caracteres que representa um intervalo.</span><span class="sxs-lookup"><span data-stu-id="0cf9d-168">The source to be copied is passed as a range or a string address representing a range.</span></span>
<span data-ttu-id="0cf9d-169">O exemplo a seguir copia dados de **A1:E1** para o intervalo que começa em **G1** (que acaba sendo colado em **G1:K1**).</span><span class="sxs-lookup"><span data-stu-id="0cf9d-169">The following code sample copies the data from **A1:E1** into the range starting at **G1** (which ends up pasting into **G1:K1**).</span></span>

```js
Excel.run(function (context) {
    var sheet = context.workbook.worksheets.getItem("Sample");
    // copy a range starting at a single cell destination
    sheet.getRange("G1").copyFrom("A1:E1");
    return context.sync();
}).catch(errorHandlerFunction);
```

<span data-ttu-id="0cf9d-170">`Range.copyFrom` tem três parâmetros opcionais.</span><span class="sxs-lookup"><span data-stu-id="0cf9d-170">`Range.copyFrom` has three optional parameters.</span></span>

```TypeScript
copyFrom(sourceRange: Range | RangeAreas | string, copyType?: Excel.RangeCopyType, skipBlanks?: boolean, transpose?: boolean): void;
```

<span data-ttu-id="0cf9d-171">`copyType` especifica quais dados são copiados da origem para o destino.</span><span class="sxs-lookup"><span data-stu-id="0cf9d-171">`copyType` specifies what data gets copied from the source to the destination.</span></span>

- <span data-ttu-id="0cf9d-172">`Excel.RangeCopyType.formulas` transfere as fórmulas nas células de origem e preserva o posicionamento relativo dos intervalos dessas fórmulas.</span><span class="sxs-lookup"><span data-stu-id="0cf9d-172">`Excel.RangeCopyType.formulas` transfers the formulas in the source cells and preserves the relative positioning of those formulas’ ranges.</span></span> <span data-ttu-id="0cf9d-173">As entradas que não sejam uma fórmula são copiadas no seu estado original.</span><span class="sxs-lookup"><span data-stu-id="0cf9d-173">Any non-formula entries are copied as-is.</span></span>
- <span data-ttu-id="0cf9d-174">`Excel.RangeCopyType.values` copia os valores dos dados e, no caso de fórmulas, o resultado da fórmula.</span><span class="sxs-lookup"><span data-stu-id="0cf9d-174">`Excel.RangeCopyType.values` copies the data values and, in the case of formulas, the result of the formula.</span></span>
- <span data-ttu-id="0cf9d-175">`Excel.RangeCopyType.formats` copia a formatação do intervalo incluindo a fonte, cor e outras configurações de formato, mas nenhum valor.</span><span class="sxs-lookup"><span data-stu-id="0cf9d-175">`Excel.RangeCopyType.formats` copies the formatting of the range, including font, color, and other format settings, but no values.</span></span>
- <span data-ttu-id="0cf9d-176">`Excel.RangeCopyType.all` (a opção padrão) copia ambos os dados e formatação, preservando as fórmulas das células, caso elas sejam encontradas.</span><span class="sxs-lookup"><span data-stu-id="0cf9d-176">`Excel.RangeCopyType.all` (the default option) copies both data and formatting, preserving cells’ formulas if found.</span></span>

<span data-ttu-id="0cf9d-177">`skipBlanks` define se as células em branco são copiadas para o destino.</span><span class="sxs-lookup"><span data-stu-id="0cf9d-177">`skipBlanks` sets whether blank cells are copied into the destination.</span></span> <span data-ttu-id="0cf9d-178">Quando for verdadeiro, `copyFrom` ignora células em branco no intervalo de origem.</span><span class="sxs-lookup"><span data-stu-id="0cf9d-178">When true, `copyFrom` skips blank cells in the source range.</span></span>
<span data-ttu-id="0cf9d-179">As células ignoradas não substituem os dados existentes de suas células correspondentes no intervalo de destino.</span><span class="sxs-lookup"><span data-stu-id="0cf9d-179">Skipped cells will not overwrite the existing data of their corresponding cells in the destination range.</span></span> <span data-ttu-id="0cf9d-180">O padrão é false.</span><span class="sxs-lookup"><span data-stu-id="0cf9d-180">The default is false.</span></span>

<span data-ttu-id="0cf9d-181">`transpose` determina se os dados são transpostos, ou seja, suas linhas e colunas são alternadas para o local de origem.</span><span class="sxs-lookup"><span data-stu-id="0cf9d-181">`transpose` determines whether or not the data is transposed, meaning its rows and columns are switched, into the source location.</span></span>
<span data-ttu-id="0cf9d-182">Um intervalo transposto invertido na diagonal principal, portanto as linhas **1**, **2** e **3** se tornarão as colunas **A**, **B** e **C**.</span><span class="sxs-lookup"><span data-stu-id="0cf9d-182">A transposed range is flipped along the main diagonal, so rows **1**, **2**, and **3** will become columns **A**, **B**, and **C**.</span></span>

<span data-ttu-id="0cf9d-183">O exemplo de código e as imagens a seguir demonstram esse comportamento em um cenário simples.</span><span class="sxs-lookup"><span data-stu-id="0cf9d-183">The following code sample and images demonstrate this behavior in a simple scenario.</span></span>

```js
Excel.run(function (context) {
    var sheet = context.workbook.worksheets.getItem("Sample");
    // copy a range, omitting the blank cells so existing data is not overwritten in those cells
    sheet.getRange("D1").copyFrom("A1:C1",
        Excel.RangeCopyType.all,
        true, // skipBlanks
        false); // transpose
    // copy a range, including the blank cells which will overwrite existing data in the target cells
    sheet.getRange("D2").copyFrom("A2:C2",
        Excel.RangeCopyType.all,
        false, // skipBlanks
        false); // transpose
    return context.sync();
}).catch(errorHandlerFunction);
```

<span data-ttu-id="0cf9d-184">*Antes da função precedente ter sido executada.*</span><span class="sxs-lookup"><span data-stu-id="0cf9d-184">*Before the preceding function has been run.*</span></span>

![Os dados no Excel antes do método de copiar do intervalo foram executados](../images/excel-range-copyfrom-skipblanks-before.png)

<span data-ttu-id="0cf9d-186">*Após a função precedente ter sido executada.*</span><span class="sxs-lookup"><span data-stu-id="0cf9d-186">*After the preceding function has been run.*</span></span>

![Os dados no Excel após o método de copiar do intervalo foram executados](../images/excel-range-copyfrom-skipblanks-after.png)

## <a name="remove-duplicates-preview"></a><span data-ttu-id="0cf9d-188">Remover duplicatas (visualização)</span><span class="sxs-lookup"><span data-stu-id="0cf9d-188">Remove duplicates (preview)</span></span>

> [!NOTE]
> <span data-ttu-id="0cf9d-189">A função `removeDuplicates` do objeto do intervalo só está disponível atualmente na versão prévia pública.</span><span class="sxs-lookup"><span data-stu-id="0cf9d-189">The Range object's `removeDuplicates` function is currently available only in public preview.</span></span> [!INCLUDE [Information about using preview APIs](../includes/using-excel-preview-apis.md)]

<span data-ttu-id="0cf9d-190">A função do objeto intervalo `removeDuplicates` remove linhas com entradas duplicadas em determinadas colunas.</span><span class="sxs-lookup"><span data-stu-id="0cf9d-190">The Range object's `removeDuplicates` function removes rows with duplicate entries in the specified columns.</span></span> <span data-ttu-id="0cf9d-191">A função passa por cada linha no intervalo do índice de menor valor até o índice de maior valor no intervalo (de cima para baixo).</span><span class="sxs-lookup"><span data-stu-id="0cf9d-191">The function goes through each row in the range from the lowest-valued index to the highest-valued index in the range (from top to bottom).</span></span> <span data-ttu-id="0cf9d-192">Uma linha é excluída se um valor em sua coluna ou colunas especificadas aparecer mais cedo no intervalo.</span><span class="sxs-lookup"><span data-stu-id="0cf9d-192">A row is deleted if a value in its specified column or columns appeared earlier in the range.</span></span> <span data-ttu-id="0cf9d-193">Linhas no intervalo abaixo da linha excluída são deslocadas para cima.</span><span class="sxs-lookup"><span data-stu-id="0cf9d-193">Rows in the range below the deleted row are shifted up.</span></span> <span data-ttu-id="0cf9d-194">`removeDuplicates` não afeta a posição de células fora do intervalo.</span><span class="sxs-lookup"><span data-stu-id="0cf9d-194">`removeDuplicates` does not affect the position of cells outside of the range.</span></span>

<span data-ttu-id="0cf9d-195">`removeDuplicates` leva um `number[]` representando os índices da coluna que são verificados para duplicatas.</span><span class="sxs-lookup"><span data-stu-id="0cf9d-195">`removeDuplicates` takes in a `number[]` representing the column indices which are checked for duplicates.</span></span> <span data-ttu-id="0cf9d-196">Essa matriz é baseada em zero e relativa ao intervalo, não à planilha.</span><span class="sxs-lookup"><span data-stu-id="0cf9d-196">This array is zero-based and relative to the range, not the worksheet.</span></span> <span data-ttu-id="0cf9d-197">A função também aceita um parâmetro booliano que especifica se a primeira linha é um cabeçalho.</span><span class="sxs-lookup"><span data-stu-id="0cf9d-197">The function also takes in a boolean parameter that specifies whether the first row is a header.</span></span> <span data-ttu-id="0cf9d-198">Quando **verdadeiro**, a primeira linha será ignorada ao considerar duplicatas.</span><span class="sxs-lookup"><span data-stu-id="0cf9d-198">When **true**, the top row is ignored when considering duplicates.</span></span> <span data-ttu-id="0cf9d-199">A função `removeDuplicates` retorna um objeto `RemoveDuplicatesResult` que especifica o número de linhas removidas e o número de linhas exclusivas restantes.</span><span class="sxs-lookup"><span data-stu-id="0cf9d-199">The `removeDuplicates` function returns a `RemoveDuplicatesResult` object that specifies the number of rows removed and the number of unique rows remaining.</span></span>

<span data-ttu-id="0cf9d-200">Ao usar um intervalo na função`removeDuplicates`, lembre-se do seguinte:</span><span class="sxs-lookup"><span data-stu-id="0cf9d-200">When using a range's `removeDuplicates` function, keep the following in mind:</span></span>

- <span data-ttu-id="0cf9d-201">`removeDuplicates` considera valores de célula, não resultados de função.</span><span class="sxs-lookup"><span data-stu-id="0cf9d-201">`removeDuplicates` considers cell values, not function results.</span></span> <span data-ttu-id="0cf9d-202">Se as duas funções diferentes forem avaliadas como o mesmo resultado, os valores de célula não são considerados duplicatas.</span><span class="sxs-lookup"><span data-stu-id="0cf9d-202">If two different functions evaluate to the same result, the cell values are not considered duplicates.</span></span>
- <span data-ttu-id="0cf9d-203">Células vazias não serão ignoradas por `removeDuplicates`.</span><span class="sxs-lookup"><span data-stu-id="0cf9d-203">Empty cells are not ignored by `removeDuplicates`.</span></span> <span data-ttu-id="0cf9d-204">O valor de uma célula vazia é tratado como qualquer outro valor.</span><span class="sxs-lookup"><span data-stu-id="0cf9d-204">The value of an empty cell is treated like any other value.</span></span> <span data-ttu-id="0cf9d-205">Isso significa que as linhas vazias contidas no intervalo serão incluídas em `RemoveDuplicatesResult`.</span><span class="sxs-lookup"><span data-stu-id="0cf9d-205">This means empty rows contained within in the range will be included in the `RemoveDuplicatesResult`.</span></span>

<span data-ttu-id="0cf9d-206">O exemplo a seguir mostra a remoção de entradas com valores duplicados na primeira coluna.</span><span class="sxs-lookup"><span data-stu-id="0cf9d-206">The following sample shows the removal of entries with duplicate values in the first column.</span></span>

```js
Excel.run(async (context) => {
    var sheet = context.workbook.worksheets.getItem("Sample");
    var range = sheet.getRange("B2:D11");

    var deleteResult = range.removeDuplicates([0],true);
    deleteResult.load();

    return context.sync().then(function () {
        console.log(deleteResult.removed + " entries with duplicate names removed.");
        console.log(deleteResult.uniqueRemaining + " entries with unique names remain in the range.");
    });
}).catch(errorHandlerFunction);
```

<span data-ttu-id="0cf9d-207">*Antes da função precedente ter sido executada.*</span><span class="sxs-lookup"><span data-stu-id="0cf9d-207">*Before the preceding function has been run.*</span></span>

![Dados no Excel antes da execução do método de remoção de duplicatas do intervalo](../images/excel-ranges-remove-duplicates-before.png)

<span data-ttu-id="0cf9d-209">*Após a função precedente ter sido executada.*</span><span class="sxs-lookup"><span data-stu-id="0cf9d-209">*After the preceding function has been run.*</span></span>

![Dados no Excel depois da execução do método de remoção de duplicatas do intervalo](../images/excel-ranges-remove-duplicates-after.png)

## <a name="see-also"></a><span data-ttu-id="0cf9d-211">Confira também</span><span class="sxs-lookup"><span data-stu-id="0cf9d-211">See also</span></span>

- [<span data-ttu-id="0cf9d-212">Trabalhar com intervalos usando a API JavaScript do Excel</span><span class="sxs-lookup"><span data-stu-id="0cf9d-212">Work with ranges using the Excel JavaScript API</span></span>](excel-add-ins-ranges.md)
- [<span data-ttu-id="0cf9d-213">Conceitos fundamentais de programação com a API JavaScript do Excel</span><span class="sxs-lookup"><span data-stu-id="0cf9d-213">Fundamental programming concepts with the Excel JavaScript API</span></span>](excel-add-ins-core-concepts.md)
- [<span data-ttu-id="0cf9d-214">Trabalhar simultaneamente com vários intervalos em suplementos do Excel</span><span class="sxs-lookup"><span data-stu-id="0cf9d-214">Work with multiple ranges simultaneously in Excel add-ins</span></span>](excel-add-ins-multiple-ranges.md)
