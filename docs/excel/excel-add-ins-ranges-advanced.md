---
title: Trabalhar com intervalos usando a API JavaScript do Excel (avançado)
description: ''
ms.date: 12/14/2018
ms.openlocfilehash: 42b1127580c46120d337553fdb86a19a78b37567
ms.sourcegitcommit: 09f124fac7b2e711e1a8be562a99624627c0699e
ms.translationtype: HT
ms.contentlocale: pt-BR
ms.lasthandoff: 12/15/2018
ms.locfileid: "27283790"
---
# <a name="work-with-ranges-using-the-excel-javascript-api-advanced"></a><span data-ttu-id="e065e-102">Trabalhar com intervalos usando a API JavaScript do Excel (avançado)</span><span class="sxs-lookup"><span data-stu-id="e065e-102">Work with ranges using the Excel JavaScript API</span></span>

<span data-ttu-id="e065e-103">Este artigo baseia-se em informações em [Trabalhar com intervalos usando a API JavaScript do Excel (fundamental)](excel-add-ins-ranges.md) fornecendo exemplos de código que mostram como executar tarefas mais avançadas com intervalos usando a API JavaScript do Excel.</span><span class="sxs-lookup"><span data-stu-id="e065e-103">This article builds upon information in [Work with ranges using the Excel JavaScript API (fundamental)](excel-add-ins-ranges.md) by providing code samples that show how to perform more advanced tasks with ranges using the Excel JavaScript API.</span></span> <span data-ttu-id="e065e-104">Para obter a lista completa de propriedades e métodos que o objeto **Range** suporta, confira [Objeto Range (API JavaScript para Excel)](https://docs.microsoft.com/javascript/api/excel/excel.range).</span><span class="sxs-lookup"><span data-stu-id="e065e-104">For the complete list of properties and methods that the **Range** object supports, see [Range Object (JavaScript API for Excel)](https://docs.microsoft.com/javascript/api/excel/excel.range).</span></span>

## <a name="work-with-dates-using-the-moment-msdate-plug-in"></a><span data-ttu-id="e065e-105">Trabalhar com datas usando o plug-in Moment-MSDate</span><span class="sxs-lookup"><span data-stu-id="e065e-105">Work with dates using the Moment-MSDate plug-in</span></span>

<span data-ttu-id="e065e-106">A [biblioteca Moment do JavaScript](https://momentjs.com/) fornece uma maneira conveniente de usar datas e carimbos de data e hora.</span><span class="sxs-lookup"><span data-stu-id="e065e-106">The [Moment JavaScript library](https://momentjs.com/) provides a convenient way to use dates and timestamps.</span></span> <span data-ttu-id="e065e-107">O [plug-in Moment-MSDate](https://www.npmjs.com/package/moment-msdate) converte o formato de momentos em um formato mais apropriado para o Excel.</span><span class="sxs-lookup"><span data-stu-id="e065e-107">The [Moment-MSDate plug-in](https://www.npmjs.com/package/moment-msdate) converts the format of moments into one preferable for Excel.</span></span> <span data-ttu-id="e065e-108">Este é o mesmo formato que a [função NOW](https://support.office.com/article/now-function-3337fd29-145a-4347-b2e6-20c904739c46) retorna.</span><span class="sxs-lookup"><span data-stu-id="e065e-108">This is the same format the [NOW function](https://support.office.com/article/now-function-3337fd29-145a-4347-b2e6-20c904739c46) returns.</span></span>

<span data-ttu-id="e065e-109">O código a seguir mostra como definir o intervalo em \*\* B4 \*\* para o carimbo de data/hora de um momento:</span><span class="sxs-lookup"><span data-stu-id="e065e-109">The following code shows how to set the range at **B4** to a moment's timestamp:</span></span>

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

<span data-ttu-id="e065e-110">É uma técnica semelhante para retirar a data da célula e convertê-la em um momento ou outro formato, conforme demonstrado no código a seguir:</span><span class="sxs-lookup"><span data-stu-id="e065e-110">It is a similar technique to get the date back out of the cell and convert it to a moment or other format, as demonstrated in the following code:</span></span>

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

<span data-ttu-id="e065e-111">Seu suplemento terá que formatar os intervalos para exibir as datas em um formato mais legível.</span><span class="sxs-lookup"><span data-stu-id="e065e-111">Your add-in will have to format the ranges to display the dates in a more human-readable form.</span></span> <span data-ttu-id="e065e-112">O exemplo de `"[$-409]m/d/yy h:mm AM/PM;@"` exibe a hora como "3/12/18 15:57".</span><span class="sxs-lookup"><span data-stu-id="e065e-112">The example of `"[$-409]m/d/yy h:mm AM/PM;@"` displays a time like "12/3/18 3:57 PM".</span></span> <span data-ttu-id="e065e-113">Para obter mais informações sobre formatos de números de data e hora, confira as "Diretrizes para formatos de data e hora" no artigo [Diretrizes de revisão para personalizar um formato de número](https://support.office.com/article/review-guidelines-for-customizing-a-number-format-c0a1d1fa-d3f4-4018-96b7-9c9354dd99f5).</span><span class="sxs-lookup"><span data-stu-id="e065e-113">For more information about date and time number formats, please see the "Guidelines for date and time formats" in the [Review guidelines for customizing a number format](https://support.office.com/article/review-guidelines-for-customizing-a-number-format-c0a1d1fa-d3f4-4018-96b7-9c9354dd99f5) article.</span></span>

## <a name="copy-and-paste"></a><span data-ttu-id="e065e-114">Copiar e colar</span><span class="sxs-lookup"><span data-stu-id="e065e-114">Copy and paste</span></span>

> [!NOTE]
> <span data-ttu-id="e065e-115">A função`Range.copyFrom` no momento só está disponível na versão prévia pública (beta).</span><span class="sxs-lookup"><span data-stu-id="e065e-115">The copyFrom function is currently available only in public preview (beta).</span></span> <span data-ttu-id="e065e-116">Para usar esse recurso, você deve usar a biblioteca beta do CDN do Office.js: https://appsforoffice.microsoft.com/lib/beta/hosted/office.js.</span><span class="sxs-lookup"><span data-stu-id="e065e-116">To use this feature, you must use the beta library of the Office.js CDN: https://appsforoffice.microsoft.com/lib/beta/hosted/office.js.</span></span>
> <span data-ttu-id="e065e-117">Se você estiver usando o TypeScript ou se seu editor de código usar arquivos de definição de tipo do TypeScript do IntelliSense, use https://appsforoffice.microsoft.com/lib/beta/hosted/office.d.ts.</span><span class="sxs-lookup"><span data-stu-id="e065e-117">If you are using TypeScript or your code editor uses TypeScript type definition files for IntelliSense, use https://appsforoffice.microsoft.com/lib/beta/hosted/office.d.ts.</span></span>

<span data-ttu-id="e065e-118">A função de `copyFrom` do intervalo replica o comportamento de copiar e colar da IU do Excel.</span><span class="sxs-lookup"><span data-stu-id="e065e-118">Range’s copyFrom function replicates the copy-and-paste behavior of the Excel UI.</span></span> <span data-ttu-id="e065e-119">O objeto de intervalo para o qual a função`copyFrom` é chamada é o destino.</span><span class="sxs-lookup"><span data-stu-id="e065e-119">The range object that copyFrom is called on is the destination.</span></span>
<span data-ttu-id="e065e-120">A fonte a ser copiada é passada como um intervalo ou um endereço de cadeia de caracteres que representa um intervalo.</span><span class="sxs-lookup"><span data-stu-id="e065e-120">The source to be copied is passed as a range or a string address representing a range.</span></span> <span data-ttu-id="e065e-121">O exemplo a seguir copia dados de **A1:E1** para o intervalo que começa em **G1** (que acaba sendo colado em **G1:K1**).</span><span class="sxs-lookup"><span data-stu-id="e065e-121">The following code sample copies the data from **A1:E1** into the range starting at **G1** (which ends up pasting into **G1:K1**).</span></span>

```js
Excel.run(function (context) {
    var sheet = context.workbook.worksheets.getItem("Sample");
    // copy a range starting at a single cell destination
    sheet.getRange("G1").copyFrom("A1:E1");
    return context.sync();
}).catch(errorHandlerFunction);
```

<span data-ttu-id="e065e-122">`Range.copyFrom` tem três parâmetros opcionais.</span><span class="sxs-lookup"><span data-stu-id="e065e-122">Range.copyFrom has three optional parameters.</span></span>

```TypeScript
copyFrom(sourceRange: Range | string, copyType?: "All" | "Formulas" | "Values" | "Formats", skipBlanks?: boolean, transpose?: boolean): void;
```

<span data-ttu-id="e065e-123">`copyType` especifica quais dados são copiados da origem para o destino.</span><span class="sxs-lookup"><span data-stu-id="e065e-123">`copyType` specifies what data gets copied from the source to the destination.</span></span>

- <span data-ttu-id="e065e-124">`"Formulas"` transfere as fórmulas nas células de origem e preserva o posicionamento relativo dos intervalos dessas fórmulas.</span><span class="sxs-lookup"><span data-stu-id="e065e-124">`"Formulas"` transfers the formulas in the source cells and preserves the relative positioning of those formulas’ ranges.</span></span> <span data-ttu-id="e065e-125">As entradas que não sejam uma fórmula são copiadas no seu estado original.</span><span class="sxs-lookup"><span data-stu-id="e065e-125">Any non-formula entries are copied as-is.</span></span>
- <span data-ttu-id="e065e-126">`"Values"` copia os valores dos dados e, no caso de fórmulas, o resultado da fórmula.</span><span class="sxs-lookup"><span data-stu-id="e065e-126">`"Values"` copies the data values and, in the case of formulas, the result of the formula.</span></span>
- <span data-ttu-id="e065e-127">`"Formats"` copia a formatação do intervalo incluindo a fonte, cor e outras configurações de formato, mas nenhum valor.</span><span class="sxs-lookup"><span data-stu-id="e065e-127">`"Formats"` copies the formatting of the range, including font, color, and other format settings, but no values.</span></span>
- <span data-ttu-id="e065e-128">`"All"` (a opção padrão) copia ambos os dados e formatação, preservando as fórmulas das células, caso elas sejam encontradas.</span><span class="sxs-lookup"><span data-stu-id="e065e-128">`"All"` (the default option) copies both data and formatting, preserving cells’ formulas if found.</span></span>

<span data-ttu-id="e065e-129">`skipBlanks` define se as células em branco são copiadas para o destino.</span><span class="sxs-lookup"><span data-stu-id="e065e-129">`skipBlanks` sets whether blank cells are copied into the destination.</span></span> <span data-ttu-id="e065e-130">Quando for verdadeiro, `copyFrom` ignora células em branco no intervalo de origem.</span><span class="sxs-lookup"><span data-stu-id="e065e-130">When true, `copyFrom` skips blank cells in the source range.</span></span>
<span data-ttu-id="e065e-131">As células ignoradas não substituem os dados existentes de suas células correspondentes no intervalo de destino.</span><span class="sxs-lookup"><span data-stu-id="e065e-131">Skipped cells will not overwrite the existing data of their corresponding cells in the destination range.</span></span> <span data-ttu-id="e065e-132">O padrão é false.</span><span class="sxs-lookup"><span data-stu-id="e065e-132">The default is false.</span></span>

<span data-ttu-id="e065e-133">`transpose` determina se os dados são transpostos, ou seja, suas linhas e colunas são alternadas para o local de origem.</span><span class="sxs-lookup"><span data-stu-id="e065e-133">`transpose` determines whether or not the data is transposed, meaning its rows and columns are switched, into the source location.</span></span>
<span data-ttu-id="e065e-134">Um intervalo transposto invertido na diagonal principal, portanto as linhas **1**, **2** e **3** se tornarão as colunas **A**, **B** e **C**.</span><span class="sxs-lookup"><span data-stu-id="e065e-134">A transposed range is flipped along the main diagonal, so rows **1**, **2**, and **3** will become columns **A**, **B**, and **C**.</span></span>

<span data-ttu-id="e065e-135">O exemplo de código e as imagens a seguir demonstram esse comportamento em um cenário simples.</span><span class="sxs-lookup"><span data-stu-id="e065e-135">The following code sample and images demonstrate this behavior in a simple scenario.</span></span>

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

<span data-ttu-id="e065e-136">*Antes da função precedente ter sido executada.*</span><span class="sxs-lookup"><span data-stu-id="e065e-136">*Before the preceeding function has been run.*</span></span>

![Os dados no Excel antes do método de copiar do intervalo foram executados](../images/excel-range-copyfrom-skipblanks-before.png)

<span data-ttu-id="e065e-138">*Após a função precedente ter sido executada.*</span><span class="sxs-lookup"><span data-stu-id="e065e-138">*After the preceeding function has been run.*</span></span>

![Os dados no Excel após o método de copiar do intervalo foram executados](../images/excel-range-copyfrom-skipblanks-after.png)

## <a name="remove-duplicates"></a><span data-ttu-id="e065e-140">Remover duplicatas</span><span class="sxs-lookup"><span data-stu-id="e065e-140">Remove duplicates</span></span>

> [!NOTE]
> <span data-ttu-id="e065e-141">A função `removeDuplicates` no momento só está disponível na versão prévia pública (beta).</span><span class="sxs-lookup"><span data-stu-id="e065e-141">The copyFrom function is currently available only in public preview (beta).</span></span> <span data-ttu-id="e065e-142">Para usar esse recurso, você deve usar a biblioteca beta do CDN do Office.js: https://appsforoffice.microsoft.com/lib/beta/hosted/office.js.</span><span class="sxs-lookup"><span data-stu-id="e065e-142">To use this feature, you must use the beta library of the Office.js CDN: https://appsforoffice.microsoft.com/lib/beta/hosted/office.js.</span></span>
> <span data-ttu-id="e065e-143">Se você estiver usando o TypeScript ou se seu editor de código usar arquivos de definição de tipo do TypeScript do IntelliSense, use https://appsforoffice.microsoft.com/lib/beta/hosted/office.d.ts.</span><span class="sxs-lookup"><span data-stu-id="e065e-143">If you are using TypeScript or your code editor uses TypeScript type definition files for IntelliSense, use https://appsforoffice.microsoft.com/lib/beta/hosted/office.d.ts.</span></span>

<span data-ttu-id="e065e-144">A função do objeto intervalo `removeDuplicates` remove linhas com entradas duplicadas em determinadas colunas.</span><span class="sxs-lookup"><span data-stu-id="e065e-144">The Range object's `removeDuplicates` function removes rows with duplicate entries in the specified columns.</span></span> <span data-ttu-id="e065e-145">A função passa por cada linha no intervalo do índice de menor valor até o índice de maior valor no intervalo (de cima para baixo).</span><span class="sxs-lookup"><span data-stu-id="e065e-145">The function goes through each row in the range from the lowest-valued index to the highest-valued index in the range (from top to bottom).</span></span> <span data-ttu-id="e065e-146">Uma linha é excluída se um valor em sua coluna ou colunas especificadas aparecer mais cedo no intervalo.</span><span class="sxs-lookup"><span data-stu-id="e065e-146">A row is deleted if a value in its specified column or columns appeared earlier in the range.</span></span> <span data-ttu-id="e065e-147">Linhas no intervalo abaixo da linha excluída são deslocadas para cima.</span><span class="sxs-lookup"><span data-stu-id="e065e-147">Rows in the range below the deleted row are shifted up.</span></span> <span data-ttu-id="e065e-148">`removeDuplicates` não afeta a posição de células fora do intervalo.</span><span class="sxs-lookup"><span data-stu-id="e065e-148">`removeDuplicates` does not affect the position of cells outside of the range.</span></span>

<span data-ttu-id="e065e-149">`removeDuplicates` leva um `number[]` representando os índices da coluna que são verificados para duplicatas.</span><span class="sxs-lookup"><span data-stu-id="e065e-149">`removeDuplicates` takes in a `number[]` representing the column indices which are checked for duplicates.</span></span> <span data-ttu-id="e065e-150">Essa matriz é baseada em zero e relativa ao intervalo, não à planilha.</span><span class="sxs-lookup"><span data-stu-id="e065e-150">This array is zero-based and relative to the range, not the worksheet.</span></span> <span data-ttu-id="e065e-151">A função também aceita um parâmetro booliano que especifica se a primeira linha é um cabeçalho.</span><span class="sxs-lookup"><span data-stu-id="e065e-151">The function also takes in a boolean parameter that specifies whether the first row is a header.</span></span> <span data-ttu-id="e065e-152">Quando **verdadeiro**, a primeira linha será ignorada ao considerar duplicatas.</span><span class="sxs-lookup"><span data-stu-id="e065e-152">When **true**, the top row is ignored when considering duplicates.</span></span> <span data-ttu-id="e065e-153">A função `removeDuplicates` retorna um objeto `RemoveDuplicatesResult` que especifica o número de linhas removidas e o número de linhas exclusivas restantes.</span><span class="sxs-lookup"><span data-stu-id="e065e-153">The `removeDuplicates` function returns a `RemoveDuplicatesResult` object that specifies the number of rows removed and the number of unique rows remaining.</span></span>

<span data-ttu-id="e065e-154">Ao usar um intervalo na função`removeDuplicates`, lembre-se do seguinte:</span><span class="sxs-lookup"><span data-stu-id="e065e-154">When using a range's `removeDuplicates` function, keep the following in mind:</span></span>

- <span data-ttu-id="e065e-155">`removeDuplicates` considera valores de célula, não resultados de função.</span><span class="sxs-lookup"><span data-stu-id="e065e-155">`removeDuplicates` considers cell values, not function results.</span></span> <span data-ttu-id="e065e-156">Se as duas funções diferentes forem avaliadas como o mesmo resultado, os valores de célula não são considerados duplicatas.</span><span class="sxs-lookup"><span data-stu-id="e065e-156">If two different functions evaluate to the same result, the cell values are not considered duplicates.</span></span>
- <span data-ttu-id="e065e-157">Células vazias não serão ignoradas por `removeDuplicates`.</span><span class="sxs-lookup"><span data-stu-id="e065e-157">Empty cells are not ignored by `removeDuplicates`.</span></span> <span data-ttu-id="e065e-158">O valor de uma célula vazia é tratado como qualquer outro valor.</span><span class="sxs-lookup"><span data-stu-id="e065e-158">The value of an empty cell is treated like any other value.</span></span> <span data-ttu-id="e065e-159">Isso significa que as linhas vazias contidas no intervalo serão incluídas em `RemoveDuplicatesResult`.</span><span class="sxs-lookup"><span data-stu-id="e065e-159">This means empty rows contained within in the range will be included in the `RemoveDuplicatesResult`.</span></span>

<span data-ttu-id="e065e-160">O exemplo a seguir mostra a remoção de entradas com valores duplicados na primeira coluna.</span><span class="sxs-lookup"><span data-stu-id="e065e-160">The following sample shows the removal of entries with duplicate values in the first column.</span></span>

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

<span data-ttu-id="e065e-161">*Antes da função precedente ter sido executada.*</span><span class="sxs-lookup"><span data-stu-id="e065e-161">*Before the preceeding function has been run.*</span></span>

![Dados no Excel antes da execução do método de remoção de duplicatas do intervalo](../images/excel-ranges-remove-duplicates-before.png)

<span data-ttu-id="e065e-163">*Após a função precedente ter sido executada.*</span><span class="sxs-lookup"><span data-stu-id="e065e-163">*After the preceeding function has been run.*</span></span>

![Dados no Excel depois da execução do método de remoção de duplicatas do intervalo](../images/excel-ranges-remove-duplicates-after.png)

## <a name="see-also"></a><span data-ttu-id="e065e-165">Confira também</span><span class="sxs-lookup"><span data-stu-id="e065e-165">See also</span></span>

- [<span data-ttu-id="e065e-166">Trabalhar com intervalos usando a API JavaScript do Excel</span><span class="sxs-lookup"><span data-stu-id="e065e-166">Work with ranges using the Excel JavaScript API</span></span>](excel-add-ins-ranges.md)
- [<span data-ttu-id="e065e-167">Conceitos fundamentais de programação com a API JavaScript do Excel</span><span class="sxs-lookup"><span data-stu-id="e065e-167">Fundamental programming concepts with the Excel JavaScript API</span></span>](excel-add-ins-core-concepts.md)