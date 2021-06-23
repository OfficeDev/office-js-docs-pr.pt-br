---
title: Intervalos de corte, cópia e colar usando a API JavaScript Excel JavaScript
description: Saiba como cortar, copiar e colar intervalos usando Excel API JavaScript.
ms.date: 04/02/2021
ms.prod: excel
localization_priority: Normal
ms.openlocfilehash: 2112702110b72e0020ed72090ce495abb3ff5366
ms.sourcegitcommit: ee9e92a968e4ad23f1e371f00d4888e4203ab772
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 06/23/2021
ms.locfileid: "53075821"
---
# <a name="cut-copy-and-paste-ranges-using-the-excel-javascript-api"></a><span data-ttu-id="8baea-103">Intervalos de corte, cópia e colar usando a API JavaScript Excel JavaScript</span><span class="sxs-lookup"><span data-stu-id="8baea-103">Cut, copy, and paste ranges using the Excel JavaScript API</span></span>

<span data-ttu-id="8baea-104">Este artigo fornece exemplos de código que cortam, copiam e colaram intervalos usando Excel API JavaScript.</span><span class="sxs-lookup"><span data-stu-id="8baea-104">This article provides code samples that cut, copy, and paste ranges using the Excel JavaScript API.</span></span> <span data-ttu-id="8baea-105">Para ver a lista completa de propriedades e métodos que o `Range` objeto oferece suporte, [consulte Excel. Classe Range](/javascript/api/excel/excel.range).</span><span class="sxs-lookup"><span data-stu-id="8baea-105">For the complete list of properties and methods that the `Range` object supports, see [Excel.Range class](/javascript/api/excel/excel.range).</span></span>

[!include[Excel cells and ranges note](../includes/note-excel-cells-and-ranges.md)]

## <a name="copy-and-paste"></a><span data-ttu-id="8baea-106">Copy and paste</span><span class="sxs-lookup"><span data-stu-id="8baea-106">Copy and paste</span></span>

<span data-ttu-id="8baea-107">O [método Range.copyFrom](/javascript/api/excel/excel.range#copyfrom-sourcerange--copytype--skipblanks--transpose-) replica as ações **Copiar** e **Colar** da interface do usuário Excel usuário.</span><span class="sxs-lookup"><span data-stu-id="8baea-107">The [Range.copyFrom](/javascript/api/excel/excel.range#copyfrom-sourcerange--copytype--skipblanks--transpose-) method replicates the **Copy** and **Paste** actions of the Excel UI.</span></span> <span data-ttu-id="8baea-108">O destino é `Range` o objeto `copyFrom` chamado.</span><span class="sxs-lookup"><span data-stu-id="8baea-108">The destination is the `Range` object that `copyFrom` is called on.</span></span> <span data-ttu-id="8baea-109">A fonte a ser copiada é passada como um intervalo ou um endereço de cadeia de caracteres que representa um intervalo.</span><span class="sxs-lookup"><span data-stu-id="8baea-109">The source to be copied is passed as a range or a string address representing a range.</span></span>

<span data-ttu-id="8baea-110">O exemplo a seguir copia dados de **A1:E1** para o intervalo que começa em **G1** (que acaba sendo colado em **G1:K1**).</span><span class="sxs-lookup"><span data-stu-id="8baea-110">The following code sample copies the data from **A1:E1** into the range starting at **G1** (which ends up pasting into **G1:K1**).</span></span>

```js
Excel.run(function (context) {
    var sheet = context.workbook.worksheets.getItem("Sample");
    // copy everything from "A1:E1" into "G1" and the cells afterwards ("G1:K1")
    sheet.getRange("G1").copyFrom("A1:E1");
    return context.sync();
}).catch(errorHandlerFunction);
```

<span data-ttu-id="8baea-111">`Range.copyFrom` tem três parâmetros opcionais.</span><span class="sxs-lookup"><span data-stu-id="8baea-111">`Range.copyFrom` has three optional parameters.</span></span>

```TypeScript
copyFrom(sourceRange: Range | RangeAreas | string, copyType?: Excel.RangeCopyType, skipBlanks?: boolean, transpose?: boolean): void;
```

<span data-ttu-id="8baea-112">`copyType` especifica quais dados são copiados da origem para o destino.</span><span class="sxs-lookup"><span data-stu-id="8baea-112">`copyType` specifies what data gets copied from the source to the destination.</span></span>

- <span data-ttu-id="8baea-113">`Excel.RangeCopyType.formulas` transfere as fórmulas nas células de origem e preserva o posicionamento relativo dos intervalos dessas fórmulas.</span><span class="sxs-lookup"><span data-stu-id="8baea-113">`Excel.RangeCopyType.formulas` transfers the formulas in the source cells and preserves the relative positioning of those formulas' ranges.</span></span> <span data-ttu-id="8baea-114">As entradas que não sejam uma fórmula são copiadas no seu estado original.</span><span class="sxs-lookup"><span data-stu-id="8baea-114">Any non-formula entries are copied as-is.</span></span>
- <span data-ttu-id="8baea-115">`Excel.RangeCopyType.values` copia os valores dos dados e, no caso de fórmulas, o resultado da fórmula.</span><span class="sxs-lookup"><span data-stu-id="8baea-115">`Excel.RangeCopyType.values` copies the data values and, in the case of formulas, the result of the formula.</span></span>
- <span data-ttu-id="8baea-116">`Excel.RangeCopyType.formats` copia a formatação do intervalo incluindo a fonte, cor e outras configurações de formato, mas nenhum valor.</span><span class="sxs-lookup"><span data-stu-id="8baea-116">`Excel.RangeCopyType.formats` copies the formatting of the range, including font, color, and other format settings, but no values.</span></span>
- <span data-ttu-id="8baea-117">`Excel.RangeCopyType.all` (a opção padrão) copia os dados e a formatação, preservando as fórmulas das células, se encontradas.</span><span class="sxs-lookup"><span data-stu-id="8baea-117">`Excel.RangeCopyType.all` (the default option) copies both data and formatting, preserving cells' formulas if found.</span></span>

<span data-ttu-id="8baea-118">`skipBlanks` define se as células em branco são copiadas para o destino.</span><span class="sxs-lookup"><span data-stu-id="8baea-118">`skipBlanks` sets whether blank cells are copied into the destination.</span></span> <span data-ttu-id="8baea-119">Quando for verdadeiro, `copyFrom` ignora células em branco no intervalo de origem.</span><span class="sxs-lookup"><span data-stu-id="8baea-119">When true, `copyFrom` skips blank cells in the source range.</span></span>
<span data-ttu-id="8baea-120">As células ignoradas não substituem os dados existentes de suas células correspondentes no intervalo de destino.</span><span class="sxs-lookup"><span data-stu-id="8baea-120">Skipped cells will not overwrite the existing data of their corresponding cells in the destination range.</span></span> <span data-ttu-id="8baea-121">O padrão é false.</span><span class="sxs-lookup"><span data-stu-id="8baea-121">The default is false.</span></span>

<span data-ttu-id="8baea-122">`transpose` determina se os dados são transpostos, ou seja, suas linhas e colunas são alternadas para o local de origem.</span><span class="sxs-lookup"><span data-stu-id="8baea-122">`transpose` determines whether or not the data is transposed, meaning its rows and columns are switched, into the source location.</span></span>
<span data-ttu-id="8baea-123">Um intervalo transposto invertido na diagonal principal, portanto as linhas **1**, **2** e **3** se tornarão as colunas **A**, **B** e **C**.</span><span class="sxs-lookup"><span data-stu-id="8baea-123">A transposed range is flipped along the main diagonal, so rows **1**, **2**, and **3** will become columns **A**, **B**, and **C**.</span></span>

<span data-ttu-id="8baea-124">O exemplo de código e as imagens a seguir demonstram esse comportamento em um cenário simples.</span><span class="sxs-lookup"><span data-stu-id="8baea-124">The following code sample and images demonstrate this behavior in a simple scenario.</span></span>

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

### <a name="data-before-range-is-copied-and-pasted"></a><span data-ttu-id="8baea-125">Dados antes que o intervalo seja copiado e passado</span><span class="sxs-lookup"><span data-stu-id="8baea-125">Data before range is copied and pasted</span></span>

![Dados em Excel antes que o método de cópia do intervalo tenha sido executado.](../images/excel-range-copyfrom-skipblanks-before.png)

### <a name="data-after-range-is-copied-and-pasted"></a><span data-ttu-id="8baea-127">Dados após o intervalo são copiados e copiados</span><span class="sxs-lookup"><span data-stu-id="8baea-127">Data after range is copied and pasted</span></span>

![Dados em Excel depois que o método de cópia do intervalo tiver sido executado.](../images/excel-range-copyfrom-skipblanks-after.png)

## <a name="cut-and-paste-move-cells"></a><span data-ttu-id="8baea-129">Cortar e colar células (mover)</span><span class="sxs-lookup"><span data-stu-id="8baea-129">Cut and paste (move) cells</span></span>

<span data-ttu-id="8baea-130">O [método Range.moveTo](/javascript/api/excel/excel.range#moveto-destinationrange-) move células para um novo local na workbook.</span><span class="sxs-lookup"><span data-stu-id="8baea-130">The [Range.moveTo](/javascript/api/excel/excel.range#moveto-destinationrange-) method moves cells to a new location in the workbook.</span></span> <span data-ttu-id="8baea-131">Esse comportamento de movimento de célula funciona [](https://support.office.com/article/Move-or-copy-cells-and-cell-contents-803d65eb-6a3e-4534-8c6f-ff12d1c4139e) da mesma forma que quando as células são movidas arrastando a borda do intervalo ou ao tomar as ações **Cortar** **e Colar.**</span><span class="sxs-lookup"><span data-stu-id="8baea-131">This cell movement behavior works the same as when cells are moved by [dragging the range border](https://support.office.com/article/Move-or-copy-cells-and-cell-contents-803d65eb-6a3e-4534-8c6f-ff12d1c4139e) or when taking the **Cut** and **Paste** actions.</span></span> <span data-ttu-id="8baea-132">Tanto a formatação quanto os valores do intervalo são movidos para o local especificado como o `destinationRange` parâmetro.</span><span class="sxs-lookup"><span data-stu-id="8baea-132">Both the formatting and values of the range are moved to the location specified as the `destinationRange` parameter.</span></span>

<span data-ttu-id="8baea-133">O exemplo de código a seguir move um intervalo com o `Range.moveTo` método.</span><span class="sxs-lookup"><span data-stu-id="8baea-133">The following code sample moves a range with the `Range.moveTo` method.</span></span> <span data-ttu-id="8baea-134">Observe que, se o intervalo de destino for menor que a fonte, ele será expandido para abranger o conteúdo de origem.</span><span class="sxs-lookup"><span data-stu-id="8baea-134">Note that if the destination range is smaller than the source, it will be expanded to encompass the source content.</span></span>

```js
Excel.run(function (context) {
    var sheet = context.workbook.worksheets.getActiveWorksheet();
    sheet.getRange("F1").values = [["Moved Range"]];

    // Move the cells "A1:E1" to "G1" (which fills the range "G1:K1").
    sheet.getRange("A1:E1").moveTo("G1");
    return context.sync();
});
```

## <a name="see-also"></a><span data-ttu-id="8baea-135">Confira também</span><span class="sxs-lookup"><span data-stu-id="8baea-135">See also</span></span>

- [<span data-ttu-id="8baea-136">Modelo de objeto JavaScript do Excel em Suplementos do Office</span><span class="sxs-lookup"><span data-stu-id="8baea-136">Excel JavaScript object model in Office Add-ins</span></span>](excel-add-ins-core-concepts.md)
- [<span data-ttu-id="8baea-137">Trabalhar com células usando a EXCEL JavaScript</span><span class="sxs-lookup"><span data-stu-id="8baea-137">Work with cells using the Excel JavaScript API</span></span>](excel-add-ins-cells.md)
- [<span data-ttu-id="8baea-138">Remover duplicatas usando a EXCEL JavaScript</span><span class="sxs-lookup"><span data-stu-id="8baea-138">Remove duplicates using the Excel JavaScript API</span></span>](excel-add-ins-ranges-remove-duplicates.md)
- [<span data-ttu-id="8baea-139">Trabalhar simultaneamente com vários intervalos em suplementos do Excel</span><span class="sxs-lookup"><span data-stu-id="8baea-139">Work with multiple ranges simultaneously in Excel add-ins</span></span>](excel-add-ins-multiple-ranges.md)
