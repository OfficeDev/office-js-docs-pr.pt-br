---
title: Trabalhar com intervalos usando a API JavaScript do Excel (avançado)
description: ''
ms.date: 09/18/2019
localization_priority: Normal
ms.openlocfilehash: d260ee6140d0153b426e530304e95025dc235b74
ms.sourcegitcommit: c8914ce0f48a0c19bbfc3276a80d090bb7ce68e1
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 09/26/2019
ms.locfileid: "37235327"
---
# <a name="work-with-ranges-using-the-excel-javascript-api-advanced"></a><span data-ttu-id="bd5ad-102">Trabalhar com intervalos usando a API JavaScript do Excel (avançado)</span><span class="sxs-lookup"><span data-stu-id="bd5ad-102">Work with ranges using the Excel JavaScript API (advanced)</span></span>

<span data-ttu-id="bd5ad-103">Este artigo baseia-se em informações em [Trabalhar com intervalos usando a API JavaScript do Excel (fundamental)](excel-add-ins-ranges.md) fornecendo exemplos de código que mostram como executar tarefas mais avançadas com intervalos usando a API JavaScript do Excel.</span><span class="sxs-lookup"><span data-stu-id="bd5ad-103">This article builds upon information in [Work with ranges using the Excel JavaScript API (fundamental)](excel-add-ins-ranges.md) by providing code samples that show how to perform more advanced tasks with ranges using the Excel JavaScript API.</span></span> <span data-ttu-id="bd5ad-104">Para obter a lista completa de propriedades e métodos que o objeto **Range** suporta, confira [Objeto Range (API JavaScript para Excel)](/javascript/api/excel/excel.range).</span><span class="sxs-lookup"><span data-stu-id="bd5ad-104">For the complete list of properties and methods that the **Range** object supports, see [Range Object (JavaScript API for Excel)](/javascript/api/excel/excel.range).</span></span>

## <a name="work-with-dates-using-the-moment-msdate-plug-in"></a><span data-ttu-id="bd5ad-105">Trabalhar com datas usando o plug-in Moment-MSDate</span><span class="sxs-lookup"><span data-stu-id="bd5ad-105">Work with dates using the Moment-MSDate plug-in</span></span>

<span data-ttu-id="bd5ad-106">A [biblioteca Moment do JavaScript](https://momentjs.com/) fornece uma maneira conveniente de usar datas e carimbos de data e hora.</span><span class="sxs-lookup"><span data-stu-id="bd5ad-106">The [Moment JavaScript library](https://momentjs.com/) provides a convenient way to use dates and timestamps.</span></span> <span data-ttu-id="bd5ad-107">O [plug-in Moment-MSDate](https://www.npmjs.com/package/moment-msdate) converte o formato de momentos em um formato mais apropriado para o Excel.</span><span class="sxs-lookup"><span data-stu-id="bd5ad-107">The [Moment-MSDate plug-in](https://www.npmjs.com/package/moment-msdate) converts the format of moments into one preferable for Excel.</span></span> <span data-ttu-id="bd5ad-108">Este é o mesmo formato que a [função NOW](https://support.office.com/article/now-function-3337fd29-145a-4347-b2e6-20c904739c46) retorna.</span><span class="sxs-lookup"><span data-stu-id="bd5ad-108">This is the same format the [NOW function](https://support.office.com/article/now-function-3337fd29-145a-4347-b2e6-20c904739c46) returns.</span></span>

<span data-ttu-id="bd5ad-109">O código a seguir mostra como definir o intervalo em \*\* B4 \*\* para o carimbo de data/hora de um momento:</span><span class="sxs-lookup"><span data-stu-id="bd5ad-109">The following code shows how to set the range at **B4** to a moment's timestamp:</span></span>

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

<span data-ttu-id="bd5ad-110">É uma técnica semelhante para retirar a data da célula e convertê-la em um momento ou outro formato, conforme demonstrado no código a seguir:</span><span class="sxs-lookup"><span data-stu-id="bd5ad-110">It is a similar technique to get the date back out of the cell and convert it to a moment or other format, as demonstrated in the following code:</span></span>

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

<span data-ttu-id="bd5ad-111">Seu suplemento terá que formatar os intervalos para exibir as datas em um formato mais legível.</span><span class="sxs-lookup"><span data-stu-id="bd5ad-111">Your add-in will have to format the ranges to display the dates in a more human-readable form.</span></span> <span data-ttu-id="bd5ad-112">O exemplo de `"[$-409]m/d/yy h:mm AM/PM;@"` exibe a hora como "3/12/18 15:57".</span><span class="sxs-lookup"><span data-stu-id="bd5ad-112">The example of `"[$-409]m/d/yy h:mm AM/PM;@"` displays a time like "12/3/18 3:57 PM".</span></span> <span data-ttu-id="bd5ad-113">Para obter mais informações sobre formatos de números de data e hora, confira as "Diretrizes para formatos de data e hora" no artigo [Diretrizes de revisão para personalizar um formato de número](https://support.office.com/article/review-guidelines-for-customizing-a-number-format-c0a1d1fa-d3f4-4018-96b7-9c9354dd99f5).</span><span class="sxs-lookup"><span data-stu-id="bd5ad-113">For more information about date and time number formats, please see the "Guidelines for date and time formats" in the [Review guidelines for customizing a number format](https://support.office.com/article/review-guidelines-for-customizing-a-number-format-c0a1d1fa-d3f4-4018-96b7-9c9354dd99f5) article.</span></span>

## <a name="work-with-multiple-ranges-simultaneously"></a><span data-ttu-id="bd5ad-114">Trabalhar com vários intervalos simultaneamente</span><span class="sxs-lookup"><span data-stu-id="bd5ad-114">Work with multiple ranges simultaneously</span></span>

<span data-ttu-id="bd5ad-115">O objeto [RangeAreas](/javascript/api/excel/excel.rangeareas) permite que o suplemento realize operações em vários intervalos de uma só vez.</span><span class="sxs-lookup"><span data-stu-id="bd5ad-115">The [RangeAreas](/javascript/api/excel/excel.rangeareas) object lets your add-in perform operations on multiple ranges at once.</span></span> <span data-ttu-id="bd5ad-116">Esses intervalos poderão ser contíguos, mas não precisam ser.</span><span class="sxs-lookup"><span data-stu-id="bd5ad-116">These ranges may be contiguous, but do not have to be.</span></span> <span data-ttu-id="bd5ad-117">`RangeAreas` são descritas ainda mais no artigo [Trabalhar com vários intervalos simultaneamente em suplementos do Excel](excel-add-ins-multiple-ranges.md).</span><span class="sxs-lookup"><span data-stu-id="bd5ad-117">`RangeAreas` are further discussed in the article [Work with multiple ranges simultaneously in Excel add-ins](excel-add-ins-multiple-ranges.md).</span></span>

## <a name="find-special-cells-within-a-range"></a><span data-ttu-id="bd5ad-118">Localizar células especiais em um intervalo</span><span class="sxs-lookup"><span data-stu-id="bd5ad-118">Find special cells within a range</span></span>

<span data-ttu-id="bd5ad-119">Os métodos [Range. getSpecialCells](/javascript/api/excel/excel.range#getspecialcells-celltype--cellvaluetype-) e [Range. getSpecialCellsOrNullObject](/javascript/api/excel/excel.range#getspecialcellsornullobject-celltype--cellvaluetype-) localizam intervalos com base nas características de suas células e nos tipos de valores de suas células.</span><span class="sxs-lookup"><span data-stu-id="bd5ad-119">The [Range.getSpecialCells](/javascript/api/excel/excel.range#getspecialcells-celltype--cellvaluetype-) and [Range.getSpecialCellsOrNullObject](/javascript/api/excel/excel.range#getspecialcellsornullobject-celltype--cellvaluetype-) methods find ranges based on the characteristics of their cells and the types of values of their cells.</span></span> <span data-ttu-id="bd5ad-120">Os dois métodos retornam `RangeAreas` objetos.</span><span class="sxs-lookup"><span data-stu-id="bd5ad-120">Both of these methods return `RangeAreas` objects.</span></span> <span data-ttu-id="bd5ad-121">Aqui estão as assinaturas dos métodos do arquivo de tipos de dados TypeScript:</span><span class="sxs-lookup"><span data-stu-id="bd5ad-121">Here are the signatures of the methods from the TypeScript data types file:</span></span>

```typescript
getSpecialCells(cellType: Excel.SpecialCellType, cellValueType?: Excel.SpecialCellValueType): Excel.RangeAreas;
```

```typescript
getSpecialCellsOrNullObject(cellType: Excel.SpecialCellType, cellValueType?: Excel.SpecialCellValueType): Excel.RangeAreas;
```

<span data-ttu-id="bd5ad-122">O exemplo a seguir usa o `getSpecialCells` método para localizar células com fórmulas.</span><span class="sxs-lookup"><span data-stu-id="bd5ad-122">The following example uses the `getSpecialCells` method to find all the cells with formulas.</span></span> <span data-ttu-id="bd5ad-123">Sobre este código, observe:</span><span class="sxs-lookup"><span data-stu-id="bd5ad-123">About this code, note:</span></span>

- <span data-ttu-id="bd5ad-124">Ele limita a parte da planilha que precisa ser pesquisada chamando primeiro `Worksheet.getUsedRange` e chamando `getSpecialCells` para apenas esse intervalo.</span><span class="sxs-lookup"><span data-stu-id="bd5ad-124">It limits the part of the sheet that needs to be searched by first calling `Worksheet.getUsedRange` and calling `getSpecialCells` for only that range.</span></span>
- <span data-ttu-id="bd5ad-125">O método `getSpecialCells` retorna um objeto `RangeAreas`, então todas as células com fórmulas serão coloridas de rosa, mesmo que não sejam todas contíguas.</span><span class="sxs-lookup"><span data-stu-id="bd5ad-125">The `getSpecialCells` method returns a `RangeAreas` object, so all of the cells with formulas will be colored pink even if they are not all contiguous.</span></span>

```js
Excel.run(function (context) {
    var sheet = context.workbook.worksheets.getActiveWorksheet();
    var usedRange = sheet.getUsedRange();
    var formulaRanges = usedRange.getSpecialCells(Excel.SpecialCellType.formulas);
    formulaRanges.format.fill.color = "pink";

    return context.sync();
})
```

<span data-ttu-id="bd5ad-126">Se nenhuma célula com característica destino existe no intervalo, `getSpecialCells` exibe um erro **ItemNotFound**.</span><span class="sxs-lookup"><span data-stu-id="bd5ad-126">If no cells with the targeted characteristic exist in the range, `getSpecialCells` throws an **ItemNotFound** error.</span></span> <span data-ttu-id="bd5ad-127">Isso desvia o fluxo de controle para um `catch` bloco, se houver um.</span><span class="sxs-lookup"><span data-stu-id="bd5ad-127">This diverts the flow of control to a `catch` block, if there is one.</span></span> <span data-ttu-id="bd5ad-128">Se não houver um `catch` bloco, o erro interromperá o método.</span><span class="sxs-lookup"><span data-stu-id="bd5ad-128">If there isn't a `catch` block, the error halts the method.</span></span>

<span data-ttu-id="bd5ad-129">Se você espera que células com característica direcionada sempre deveriam existir, provavelmente desejará o código para gerar um erro se as células não estiverem lá.</span><span class="sxs-lookup"><span data-stu-id="bd5ad-129">If you expect that cells with the targeted characteristic should always exist, you'll likely want your code to throw an error if those cells aren't there.</span></span> <span data-ttu-id="bd5ad-130">Se for um cenário válido que não há uma ou mais células correspondentes, o código deve verificar se há essa possibilidade e tratar normalmente sem enviar um erro.</span><span class="sxs-lookup"><span data-stu-id="bd5ad-130">If it's a valid scenario that there aren't any matching cells, your code should check for this possibility and handle it gracefully without throwing an error.</span></span> <span data-ttu-id="bd5ad-131">Você pode obter esse comportamento com o `getSpecialCellsOrNullObject` método e sua propriedade retornada `isNullObject`.</span><span class="sxs-lookup"><span data-stu-id="bd5ad-131">You can achieve this behavior with the `getSpecialCellsOrNullObject` method and its returned `isNullObject` property.</span></span> <span data-ttu-id="bd5ad-132">O exemplo a seguir usa esse padrão.</span><span class="sxs-lookup"><span data-stu-id="bd5ad-132">The following example uses this pattern.</span></span> <span data-ttu-id="bd5ad-133">Sobre este código, observe:</span><span class="sxs-lookup"><span data-stu-id="bd5ad-133">About this code, note:</span></span>

- <span data-ttu-id="bd5ad-134">O método `getSpecialCellsOrNullObject` sempre retorna um objeto de proxy, portanto, `null` nunca está no sentido comum do JavaScript.</span><span class="sxs-lookup"><span data-stu-id="bd5ad-134">The `getSpecialCellsOrNullObject` method always returns a proxy object, so it is never `null` in the ordinary JavaScript sense.</span></span> <span data-ttu-id="bd5ad-135">Mas se nenhuma célula de correspondência for encontrada, as propriedades do objeto`isNullObject` serão definida como `true`.</span><span class="sxs-lookup"><span data-stu-id="bd5ad-135">But if no matching cells are found, the `isNullObject` property of the object is set to `true`.</span></span>
- <span data-ttu-id="bd5ad-136">Ele chama `context.sync` *antes* de testar a propriedade`isNullObject`.</span><span class="sxs-lookup"><span data-stu-id="bd5ad-136">It calls `context.sync` *before* it tests the `isNullObject` property.</span></span> <span data-ttu-id="bd5ad-137">Esse é um requisito com todos os métodos e propriedades `*OrNullObject`, pois sempre terá que carregar e sincronizar as propriedades na ordem para lê-la.</span><span class="sxs-lookup"><span data-stu-id="bd5ad-137">This is a requirement with all `*OrNullObject` methods and properties, because you always have to load and sync a property in order to read it.</span></span> <span data-ttu-id="bd5ad-138">No entanto, não é necessário carregar *explicitamente* a propriedade`isNullObject`.</span><span class="sxs-lookup"><span data-stu-id="bd5ad-138">However, it is not necessary to *explicitly* load the `isNullObject` property.</span></span> <span data-ttu-id="bd5ad-139">Será carregado automaticamente pelo `context.sync` mesmo se `load` não for chamado no objeto.</span><span class="sxs-lookup"><span data-stu-id="bd5ad-139">It is automatically loaded by the `context.sync` even if `load` is not called on the object.</span></span> <span data-ttu-id="bd5ad-140">Para saber mais, confira [ \*OrNullObject](/office/dev/add-ins/excel/excel-add-ins-advanced-concepts#ornullobject-methods).</span><span class="sxs-lookup"><span data-stu-id="bd5ad-140">For more information, see [\*OrNullObject](/office/dev/add-ins/excel/excel-add-ins-advanced-concepts#ornullobject-methods).</span></span>
- <span data-ttu-id="bd5ad-141">Você pode testar esse código selecionando primeiro um intervalo sem células de fórmula e executando-o.</span><span class="sxs-lookup"><span data-stu-id="bd5ad-141">You can test this code by first selecting a range that has no formula cells and running it.</span></span> <span data-ttu-id="bd5ad-142">Selecione um intervalo que tem pelo menos uma célula com uma fórmula e execute novamente.</span><span class="sxs-lookup"><span data-stu-id="bd5ad-142">Then select a range that has at least one cell with a formula and run it again.</span></span>

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

<span data-ttu-id="bd5ad-143">Para manter a simplicidade, todos os outros exemplos deste artigo usam o método `getSpecialCells` em vez de `getSpecialCellsOrNullObject`.</span><span class="sxs-lookup"><span data-stu-id="bd5ad-143">For simplicity, all other examples in this article use the `getSpecialCells` method instead of  `getSpecialCellsOrNullObject`.</span></span>

### <a name="narrow-the-target-cells-with-cell-value-types"></a><span data-ttu-id="bd5ad-144">Restrinja as células de destino com tipos de valor de célula</span><span class="sxs-lookup"><span data-stu-id="bd5ad-144">Narrow the target cells with cell value types</span></span>

<span data-ttu-id="bd5ad-145">As `Range.getSpecialCells()` e `Range.getSpecialCellsOrNullObject()` métodos aceitam um segundo parâmetro opcional usado para restringir ainda mais as células de destino.</span><span class="sxs-lookup"><span data-stu-id="bd5ad-145">The `Range.getSpecialCells()` and `Range.getSpecialCellsOrNullObject()` methods accept an optional second parameter used to further narrow down the targeted cells.</span></span> <span data-ttu-id="bd5ad-146">Este segundo parâmetro é uma `Excel.SpecialCellValueType` você usar para especificar que você quer apenas células que contêm determinados tipos de valores.</span><span class="sxs-lookup"><span data-stu-id="bd5ad-146">This second parameter is an `Excel.SpecialCellValueType` you use to specify that you only want cells that contain certain types of values.</span></span>

> [!NOTE]
> <span data-ttu-id="bd5ad-147">O `Excel.SpecialCellValueType` parâmetro só pode ser usado se a `Excel.SpecialCellType` está `Excel.SpecialCellType.formulas` ou `Excel.SpecialCellType.constants`.</span><span class="sxs-lookup"><span data-stu-id="bd5ad-147">The `Excel.SpecialCellValueType` parameter can only be used if the `Excel.SpecialCellType` is `Excel.SpecialCellType.formulas` or `Excel.SpecialCellType.constants`.</span></span>

#### <a name="test-for-a-single-cell-value-type"></a><span data-ttu-id="bd5ad-148">Teste para um tipo de valor da célula única</span><span class="sxs-lookup"><span data-stu-id="bd5ad-148">Test for a single cell value type</span></span>

<span data-ttu-id="bd5ad-149">O `Excel.SpecialCellValueType` enumeração com esses quatro tipos básicos (além dos outros valores combinados descritos nesta seção posterior):</span><span class="sxs-lookup"><span data-stu-id="bd5ad-149">The `Excel.SpecialCellValueType` enum has these four basic types (in addition to the other combined values described later in this section):</span></span>

- `Excel.SpecialCellValueType.errors`
- <span data-ttu-id="bd5ad-150">`Excel.SpecialCellValueType.logical` (ou seja, booliano)</span><span class="sxs-lookup"><span data-stu-id="bd5ad-150">`Excel.SpecialCellValueType.logical` (which means boolean)</span></span>
- `Excel.SpecialCellValueType.numbers`
- `Excel.SpecialCellValueType.text`

<span data-ttu-id="bd5ad-151">O exemplo a seguir localiza as células especiais que são constantes numéricos e colore essas células em rosa.</span><span class="sxs-lookup"><span data-stu-id="bd5ad-151">The following example finds special cells that are numerical constants and colors those cells pink.</span></span> <span data-ttu-id="bd5ad-152">Sobre este código, observe:</span><span class="sxs-lookup"><span data-stu-id="bd5ad-152">About this code, note:</span></span>

- <span data-ttu-id="bd5ad-153">Ele apenas irá realçar células que contêm um valor numérico literal.</span><span class="sxs-lookup"><span data-stu-id="bd5ad-153">It will only highlight cells that have a literal number value.</span></span> <span data-ttu-id="bd5ad-154">Ele não destacará as células que têm uma fórmula (mesmo se o resultado for um número) ou células de estado booliano, de texto ou de erro.</span><span class="sxs-lookup"><span data-stu-id="bd5ad-154">It will not highlight cells that have a formula (even if the result is a number) or a boolean, text, or error state cells.</span></span>
- <span data-ttu-id="bd5ad-155">Para testar o código, certifique-se de que a planilha tenha algumas células com valores numéricos literais, algumas com outros tipos de valores literais e algumas com fórmulas.</span><span class="sxs-lookup"><span data-stu-id="bd5ad-155">To test the code, be sure the worksheet has some cells with literal number values, some with other kinds of literal values, and some with formulas.</span></span>

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

#### <a name="test-for-multiple-cell-value-types"></a><span data-ttu-id="bd5ad-156">Teste para vários tipos de valores de célula</span><span class="sxs-lookup"><span data-stu-id="bd5ad-156">Test for multiple cell value types</span></span>

<span data-ttu-id="bd5ad-157">Às vezes, você precisa operar com mais de um tipo de valor de célula, como todas as células com valor de texto e com valor booliano ("Lógico"). (`Excel.SpecialCellValueType.logical`).</span><span class="sxs-lookup"><span data-stu-id="bd5ad-157">Sometimes you need to operate on more than one cell value type, such as all text-valued and all boolean-valued (`Excel.SpecialCellValueType.logical`) cells.</span></span> <span data-ttu-id="bd5ad-158">O `Excel.SpecialCellValueType` enumeração tem valores com tipos combinado.</span><span class="sxs-lookup"><span data-stu-id="bd5ad-158">The `Excel.SpecialCellValueType` enum has values with combined types.</span></span> <span data-ttu-id="bd5ad-159">Por exemplo, `Excel.SpecialCellValueType.logicalText` segmentará todas as células boolianas e todos os valores de texto.</span><span class="sxs-lookup"><span data-stu-id="bd5ad-159">For example, `Excel.SpecialCellValueType.logicalText` targets all boolean and all text-valued cells.</span></span> <span data-ttu-id="bd5ad-160">`Excel.SpecialCellValueType.all` é o valor padrão, que não limita os tipos de valor da célula retornados.</span><span class="sxs-lookup"><span data-stu-id="bd5ad-160">`Excel.SpecialCellValueType.all` is the default value, which does not limit the cell value types returned.</span></span> <span data-ttu-id="bd5ad-161">O exemplo a seguir destaca todas as células com fórmulas que produzem valores ou números boolianos.</span><span class="sxs-lookup"><span data-stu-id="bd5ad-161">The following example colors all cells with formulas that produce number or boolean value.</span></span>

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

## <a name="copy-and-paste"></a><span data-ttu-id="bd5ad-162">Copy and paste</span><span class="sxs-lookup"><span data-stu-id="bd5ad-162">Copy and paste</span></span>

<span data-ttu-id="bd5ad-163">O método [Range. copyFrom](/javascript/api/excel/excel.range#copyfrom-sourcerange--copytype--skipblanks--transpose-) Replica o comportamento de copiar e colar da interface do usuário do Excel.</span><span class="sxs-lookup"><span data-stu-id="bd5ad-163">The [Range.copyFrom](/javascript/api/excel/excel.range#copyfrom-sourcerange--copytype--skipblanks--transpose-) method replicates the copy-and-paste behavior of the Excel UI.</span></span> <span data-ttu-id="bd5ad-164">O objeto de intervalo para o qual a função`copyFrom` é chamada é o destino.</span><span class="sxs-lookup"><span data-stu-id="bd5ad-164">The range object that `copyFrom` is called on is the destination.</span></span> <span data-ttu-id="bd5ad-165">A fonte a ser copiada é passada como um intervalo ou um endereço de cadeia de caracteres que representa um intervalo.</span><span class="sxs-lookup"><span data-stu-id="bd5ad-165">The source to be copied is passed as a range or a string address representing a range.</span></span>

<span data-ttu-id="bd5ad-166">O exemplo a seguir copia dados de **A1:E1** para o intervalo que começa em **G1** (que acaba sendo colado em **G1:K1**).</span><span class="sxs-lookup"><span data-stu-id="bd5ad-166">The following code sample copies the data from **A1:E1** into the range starting at **G1** (which ends up pasting into **G1:K1**).</span></span>

```js
Excel.run(function (context) {
    var sheet = context.workbook.worksheets.getItem("Sample");
    // copy everything from "A1:E1" into "G1" and the cells afterwards ("G1:K1")
    sheet.getRange("G1").copyFrom("A1:E1");
    return context.sync();
}).catch(errorHandlerFunction);
```

<span data-ttu-id="bd5ad-167">`Range.copyFrom` tem três parâmetros opcionais.</span><span class="sxs-lookup"><span data-stu-id="bd5ad-167">`Range.copyFrom` has three optional parameters.</span></span>

```TypeScript
copyFrom(sourceRange: Range | RangeAreas | string, copyType?: Excel.RangeCopyType, skipBlanks?: boolean, transpose?: boolean): void;
```

<span data-ttu-id="bd5ad-168">`copyType` especifica quais dados são copiados da origem para o destino.</span><span class="sxs-lookup"><span data-stu-id="bd5ad-168">`copyType` specifies what data gets copied from the source to the destination.</span></span>

- <span data-ttu-id="bd5ad-169">`Excel.RangeCopyType.formulas` transfere as fórmulas nas células de origem e preserva o posicionamento relativo dos intervalos dessas fórmulas.</span><span class="sxs-lookup"><span data-stu-id="bd5ad-169">`Excel.RangeCopyType.formulas` transfers the formulas in the source cells and preserves the relative positioning of those formulas’ ranges.</span></span> <span data-ttu-id="bd5ad-170">As entradas que não sejam uma fórmula são copiadas no seu estado original.</span><span class="sxs-lookup"><span data-stu-id="bd5ad-170">Any non-formula entries are copied as-is.</span></span>
- <span data-ttu-id="bd5ad-171">`Excel.RangeCopyType.values` copia os valores dos dados e, no caso de fórmulas, o resultado da fórmula.</span><span class="sxs-lookup"><span data-stu-id="bd5ad-171">`Excel.RangeCopyType.values` copies the data values and, in the case of formulas, the result of the formula.</span></span>
- <span data-ttu-id="bd5ad-172">`Excel.RangeCopyType.formats` copia a formatação do intervalo incluindo a fonte, cor e outras configurações de formato, mas nenhum valor.</span><span class="sxs-lookup"><span data-stu-id="bd5ad-172">`Excel.RangeCopyType.formats` copies the formatting of the range, including font, color, and other format settings, but no values.</span></span>
- <span data-ttu-id="bd5ad-173">`Excel.RangeCopyType.all` (a opção padrão) copia ambos os dados e formatação, preservando as fórmulas das células, caso elas sejam encontradas.</span><span class="sxs-lookup"><span data-stu-id="bd5ad-173">`Excel.RangeCopyType.all` (the default option) copies both data and formatting, preserving cells’ formulas if found.</span></span>

<span data-ttu-id="bd5ad-174">`skipBlanks` define se as células em branco são copiadas para o destino.</span><span class="sxs-lookup"><span data-stu-id="bd5ad-174">`skipBlanks` sets whether blank cells are copied into the destination.</span></span> <span data-ttu-id="bd5ad-175">Quando for verdadeiro, `copyFrom` ignora células em branco no intervalo de origem.</span><span class="sxs-lookup"><span data-stu-id="bd5ad-175">When true, `copyFrom` skips blank cells in the source range.</span></span>
<span data-ttu-id="bd5ad-176">As células ignoradas não substituem os dados existentes de suas células correspondentes no intervalo de destino.</span><span class="sxs-lookup"><span data-stu-id="bd5ad-176">Skipped cells will not overwrite the existing data of their corresponding cells in the destination range.</span></span> <span data-ttu-id="bd5ad-177">O padrão é false.</span><span class="sxs-lookup"><span data-stu-id="bd5ad-177">The default is false.</span></span>

<span data-ttu-id="bd5ad-178">`transpose` determina se os dados são transpostos, ou seja, suas linhas e colunas são alternadas para o local de origem.</span><span class="sxs-lookup"><span data-stu-id="bd5ad-178">`transpose` determines whether or not the data is transposed, meaning its rows and columns are switched, into the source location.</span></span>
<span data-ttu-id="bd5ad-179">Um intervalo transposto invertido na diagonal principal, portanto as linhas **1**, **2** e **3** se tornarão as colunas **A**, **B** e **C**.</span><span class="sxs-lookup"><span data-stu-id="bd5ad-179">A transposed range is flipped along the main diagonal, so rows **1**, **2**, and **3** will become columns **A**, **B**, and **C**.</span></span>

<span data-ttu-id="bd5ad-180">O exemplo de código e as imagens a seguir demonstram esse comportamento em um cenário simples.</span><span class="sxs-lookup"><span data-stu-id="bd5ad-180">The following code sample and images demonstrate this behavior in a simple scenario.</span></span>

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

<span data-ttu-id="bd5ad-181">*Antes da função precedente ter sido executada.*</span><span class="sxs-lookup"><span data-stu-id="bd5ad-181">*Before the preceding function has been run.*</span></span>

![Os dados no Excel antes do método de copiar do intervalo foram executados](../images/excel-range-copyfrom-skipblanks-before.png)

<span data-ttu-id="bd5ad-183">*Após a função precedente ter sido executada.*</span><span class="sxs-lookup"><span data-stu-id="bd5ad-183">*After the preceding function has been run.*</span></span>

![Os dados no Excel após o método de copiar do intervalo foram executados](../images/excel-range-copyfrom-skipblanks-after.png)

## <a name="remove-duplicates"></a><span data-ttu-id="bd5ad-185">Remover duplicatas</span><span class="sxs-lookup"><span data-stu-id="bd5ad-185">Remove duplicates</span></span>

<span data-ttu-id="bd5ad-186">O método [Range. removeDuplicates](/javascript/api/excel/excel.range#removeduplicates-columns--includesheader-) remove linhas com entradas duplicadas nas colunas especificadas.</span><span class="sxs-lookup"><span data-stu-id="bd5ad-186">The [Range.removeDuplicates](/javascript/api/excel/excel.range#removeduplicates-columns--includesheader-) method removes rows with duplicate entries in the specified columns.</span></span> <span data-ttu-id="bd5ad-187">O método passa por todas as linhas no intervalo do índice de valor mais baixo para o índice de valor mais alto no intervalo (de cima para baixo).</span><span class="sxs-lookup"><span data-stu-id="bd5ad-187">The method goes through each row in the range from the lowest-valued index to the highest-valued index in the range (from top to bottom).</span></span> <span data-ttu-id="bd5ad-188">Uma linha é excluída se um valor em sua coluna ou colunas especificadas aparecer mais cedo no intervalo.</span><span class="sxs-lookup"><span data-stu-id="bd5ad-188">A row is deleted if a value in its specified column or columns appeared earlier in the range.</span></span> <span data-ttu-id="bd5ad-189">Linhas no intervalo abaixo da linha excluída são deslocadas para cima.</span><span class="sxs-lookup"><span data-stu-id="bd5ad-189">Rows in the range below the deleted row are shifted up.</span></span> <span data-ttu-id="bd5ad-190">`removeDuplicates` não afeta a posição de células fora do intervalo.</span><span class="sxs-lookup"><span data-stu-id="bd5ad-190">`removeDuplicates` does not affect the position of cells outside of the range.</span></span>

<span data-ttu-id="bd5ad-191">`removeDuplicates` leva um `number[]` representando os índices da coluna que são verificados para duplicatas.</span><span class="sxs-lookup"><span data-stu-id="bd5ad-191">`removeDuplicates` takes in a `number[]` representing the column indices which are checked for duplicates.</span></span> <span data-ttu-id="bd5ad-192">Essa matriz é baseada em zero e relativa ao intervalo, não à planilha.</span><span class="sxs-lookup"><span data-stu-id="bd5ad-192">This array is zero-based and relative to the range, not the worksheet.</span></span> <span data-ttu-id="bd5ad-193">O método também utiliza um parâmetro Boolean que especifica se a primeira linha é um cabeçalho.</span><span class="sxs-lookup"><span data-stu-id="bd5ad-193">The method also takes in a boolean parameter that specifies whether the first row is a header.</span></span> <span data-ttu-id="bd5ad-194">Quando **verdadeiro**, a primeira linha será ignorada ao considerar duplicatas.</span><span class="sxs-lookup"><span data-stu-id="bd5ad-194">When **true**, the top row is ignored when considering duplicates.</span></span> <span data-ttu-id="bd5ad-195">O `removeDuplicates` método retorna um `RemoveDuplicatesResult` objeto que especifica o número de linhas removidas e o número de linhas exclusivas restantes.</span><span class="sxs-lookup"><span data-stu-id="bd5ad-195">The `removeDuplicates` method returns a `RemoveDuplicatesResult` object that specifies the number of rows removed and the number of unique rows remaining.</span></span>

<span data-ttu-id="bd5ad-196">Ao usar o método de `removeDuplicates` um intervalo, lembre-se do seguinte:</span><span class="sxs-lookup"><span data-stu-id="bd5ad-196">When using a range's `removeDuplicates` method, keep the following in mind:</span></span>

- <span data-ttu-id="bd5ad-197">`removeDuplicates` considera valores de célula, não resultados de função.</span><span class="sxs-lookup"><span data-stu-id="bd5ad-197">`removeDuplicates` considers cell values, not function results.</span></span> <span data-ttu-id="bd5ad-198">Se as duas funções diferentes forem avaliadas como o mesmo resultado, os valores de célula não são considerados duplicatas.</span><span class="sxs-lookup"><span data-stu-id="bd5ad-198">If two different functions evaluate to the same result, the cell values are not considered duplicates.</span></span>
- <span data-ttu-id="bd5ad-199">Células vazias não serão ignoradas por `removeDuplicates`.</span><span class="sxs-lookup"><span data-stu-id="bd5ad-199">Empty cells are not ignored by `removeDuplicates`.</span></span> <span data-ttu-id="bd5ad-200">O valor de uma célula vazia é tratado como qualquer outro valor.</span><span class="sxs-lookup"><span data-stu-id="bd5ad-200">The value of an empty cell is treated like any other value.</span></span> <span data-ttu-id="bd5ad-201">Isso significa que as linhas vazias contidas no intervalo serão incluídas em `RemoveDuplicatesResult`.</span><span class="sxs-lookup"><span data-stu-id="bd5ad-201">This means empty rows contained within in the range will be included in the `RemoveDuplicatesResult`.</span></span>

<span data-ttu-id="bd5ad-202">O exemplo a seguir mostra a remoção de entradas com valores duplicados na primeira coluna.</span><span class="sxs-lookup"><span data-stu-id="bd5ad-202">The following sample shows the removal of entries with duplicate values in the first column.</span></span>

```js
Excel.run(function (context) {
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

<span data-ttu-id="bd5ad-203">*Antes da função precedente ter sido executada.*</span><span class="sxs-lookup"><span data-stu-id="bd5ad-203">*Before the preceding function has been run.*</span></span>

![Dados no Excel antes da execução do método de remoção de duplicatas do intervalo](../images/excel-ranges-remove-duplicates-before.png)

<span data-ttu-id="bd5ad-205">*Após a função precedente ter sido executada.*</span><span class="sxs-lookup"><span data-stu-id="bd5ad-205">*After the preceding function has been run.*</span></span>

![Dados no Excel depois da execução do método de remoção de duplicatas do intervalo](../images/excel-ranges-remove-duplicates-after.png)

## <a name="group-data-for-an-outline"></a><span data-ttu-id="bd5ad-207">Agrupar dados para uma estrutura de tópicos</span><span class="sxs-lookup"><span data-stu-id="bd5ad-207">Group data for an outline</span></span>

> [!NOTE]
> <span data-ttu-id="bd5ad-208">As APIs de estrutura de tópicos para agrupar linhas e colunas atualmente estão disponíveis somente na visualização pública.</span><span class="sxs-lookup"><span data-stu-id="bd5ad-208">The outline APIs for grouping rows and columns are currently available only in public preview.</span></span> [!INCLUDE [Information about using preview APIs](../includes/using-excel-preview-apis.md)]

<span data-ttu-id="bd5ad-209">As linhas ou colunas de um intervalo podem ser agrupadas para criar uma [estrutura de tópicos](https://support.office.com/article/Outline-group-data-in-a-worksheet-08CE98C4-0063-4D42-8AC7-8278C49E9AFF).</span><span class="sxs-lookup"><span data-stu-id="bd5ad-209">Rows or columns of a range can be grouped together to create an [outline](https://support.office.com/article/Outline-group-data-in-a-worksheet-08CE98C4-0063-4D42-8AC7-8278C49E9AFF).</span></span> <span data-ttu-id="bd5ad-210">Esses grupos podem ser recolhidos e expandidos para ocultar e mostrar as células correspondentes.</span><span class="sxs-lookup"><span data-stu-id="bd5ad-210">These groups can be collapsed and expanded to hide and show the corresponding cells.</span></span> <span data-ttu-id="bd5ad-211">Isso facilita a análise rápida dos dados de linha principal.</span><span class="sxs-lookup"><span data-stu-id="bd5ad-211">This makes quick analysis of top-line data easier.</span></span> <span data-ttu-id="bd5ad-212">Use [Range. Group](/javascript/api/excel/excel.range#group-groupoption-) para tornar esses grupos de estrutura de tópicos.</span><span class="sxs-lookup"><span data-stu-id="bd5ad-212">Use [Range.group](/javascript/api/excel/excel.range#group-groupoption-) to make these outline groups.</span></span>

<span data-ttu-id="bd5ad-213">Uma estrutura de tópicos pode ter uma hierarquia, onde grupos menores estão aninhados em grupos maiores.</span><span class="sxs-lookup"><span data-stu-id="bd5ad-213">An outline can have a hierarchy, where smaller groups are nested under larger groups.</span></span> <span data-ttu-id="bd5ad-214">Isso permite que a estrutura de tópicos seja exibida em diferentes níveis.</span><span class="sxs-lookup"><span data-stu-id="bd5ad-214">This allows the outline to be viewed at different levels.</span></span> <span data-ttu-id="bd5ad-215">Alterar o nível de estrutura de tópicos visível pode ser feito programaticamente por meio do método [Range. showOutlineLevels](/javascript/api/excel/excel.range#showOutlineLevels-rowLevels--columnLevels-) .</span><span class="sxs-lookup"><span data-stu-id="bd5ad-215">Changing the visible outline level can be done programmatically through the [Range.showOutlineLevels](/javascript/api/excel/excel.range#showOutlineLevels-rowLevels--columnLevels-) method.</span></span> <span data-ttu-id="bd5ad-216">Observe que o Excel só oferece suporte a oito níveis de grupos de estrutura de tópicos.</span><span class="sxs-lookup"><span data-stu-id="bd5ad-216">Note that Excel only supports eight levels of outline groups.</span></span>

<span data-ttu-id="bd5ad-217">O exemplo de código a seguir mostra como criar uma estrutura de tópicos com dois níveis de grupos para ambas as linhas e colunas.</span><span class="sxs-lookup"><span data-stu-id="bd5ad-217">The following code sample shows how to create an outline with two levels of groups for both the rows and columns.</span></span> <span data-ttu-id="bd5ad-218">A imagem subsequente mostra os agrupamentos dessa estrutura de tópicos.</span><span class="sxs-lookup"><span data-stu-id="bd5ad-218">The subsequent image shows the groupings of that outline.</span></span> <span data-ttu-id="bd5ad-219">Observe que, no exemplo de código, os intervalos que estão sendo agrupados não incluem a linha ou coluna do controle de estrutura de tópicos (o "total" para este exemplo).</span><span class="sxs-lookup"><span data-stu-id="bd5ad-219">Note that in the code sample, the ranges being grouped do not include the row or column of the outline control (the "Totals" for this example).</span></span> <span data-ttu-id="bd5ad-220">Um grupo define o que será recolhido, não a linha ou coluna com o controle.</span><span class="sxs-lookup"><span data-stu-id="bd5ad-220">A group defines what will be collapsed, not the row or column with the control.</span></span>

```js
Excel.run(function (context) {
    var sheet = context.workbook.worksheets.getItem("Sample");

    // Group the larger, main level. Note that the outline controls
    // will be on row 10, meaning 4-9 will collapse and expand.
    sheet.getRange("4:9").group(Excel.GroupOption.byRows);

    // Group the smaller, sublevels. Note that the outline controls
    // will be on rows 6 and 9, meaning 4-5 and 7-8 will collapse and expand.
    sheet.getRange("4:5").group(Excel.GroupOption.byRows);
    sheet.getRange("7:8").group(Excel.GroupOption.byRows);

    // Group the larger, main level. Note that the outline controls
    // will be on column R, meaning C-Q will collapse and expand.
    sheet.getRange("C:Q").group(Excel.GroupOption.byColumns);

    // Group the smaller, sublevels. Note that the outline controls
    // will be on columns G, L, and R, meaning C-F, H-K, and M-P will collapse and expand.
    sheet.getRange("C:F").group(Excel.GroupOption.byColumns);
    sheet.getRange("H:K").group(Excel.GroupOption.byColumns);
    sheet.getRange("M:P").group(Excel.GroupOption.byColumns);
    return context.sync();
}).catch(errorHandlerFunction);

```

![Um intervalo com um contorno de duas dimensões de dois níveis](../images/excel-outline.png)

<span data-ttu-id="bd5ad-222">Para desagrupar um grupo de linhas ou colunas, use o método [Range. Upgroup](/javascript/api/excel/excel.range#ungroup-groupoption-) .</span><span class="sxs-lookup"><span data-stu-id="bd5ad-222">To ungroup a row or column group, use the [Range.upgroup](/javascript/api/excel/excel.range#ungroup-groupoption-) method.</span></span> <span data-ttu-id="bd5ad-223">Isso remove o nível mais externo da estrutura de tópicos.</span><span class="sxs-lookup"><span data-stu-id="bd5ad-223">This removes the outermost level from the outline.</span></span> <span data-ttu-id="bd5ad-224">Se vários grupos do mesmo tipo de linha ou coluna estiverem no mesmo nível no intervalo especificado, todos esses grupos serão desagrupados.</span><span class="sxs-lookup"><span data-stu-id="bd5ad-224">If multiple groups of the same row or column type are at the same level within the specified range, all of those groups are ungrouped.</span></span>

## <a name="see-also"></a><span data-ttu-id="bd5ad-225">Confira também</span><span class="sxs-lookup"><span data-stu-id="bd5ad-225">See also</span></span>

- [<span data-ttu-id="bd5ad-226">Trabalhar com intervalos usando a API JavaScript do Excel</span><span class="sxs-lookup"><span data-stu-id="bd5ad-226">Work with ranges using the Excel JavaScript API</span></span>](excel-add-ins-ranges.md)
- [<span data-ttu-id="bd5ad-227">Conceitos fundamentais de programação com a API JavaScript do Excel</span><span class="sxs-lookup"><span data-stu-id="bd5ad-227">Fundamental programming concepts with the Excel JavaScript API</span></span>](excel-add-ins-core-concepts.md)
- [<span data-ttu-id="bd5ad-228">Trabalhar simultaneamente com vários intervalos em suplementos do Excel</span><span class="sxs-lookup"><span data-stu-id="bd5ad-228">Work with multiple ranges simultaneously in Excel add-ins</span></span>](excel-add-ins-multiple-ranges.md)
