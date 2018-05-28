---
title: Adicionar valida??o de dados a intervalos do Excel
description: ''
ms.date: 04/13/2018
ms.openlocfilehash: 8e5f09f1c566103f34ad584885769229c17ab1f7
ms.sourcegitcommit: c72c35e8389c47a795afbac1b2bcf98c8e216d82
ms.translationtype: HT
ms.contentlocale: pt-BR
ms.lasthandoff: 05/23/2018
---
# <a name="add-data-validation-to-excel-ranges-preview"></a><span data-ttu-id="8a164-102">Adicionar valida??o de dados a intervalos do Excel (vers?o pr?via)</span><span class="sxs-lookup"><span data-stu-id="8a164-102">Add data validation to Excel ranges (Preview)</span></span>

> [!NOTE]
> <span data-ttu-id="8a164-103">Enquanto as APIs de valida??o de dados est?o em vers?o pr?via, voc? deve carregar a vers?o beta da biblioteca JavaScript do Office para us?-las.</span><span class="sxs-lookup"><span data-stu-id="8a164-103">While the data validation APIs are in preview, you must load the beta version of the Office JavaScript library to use them.</span></span> <span data-ttu-id="8a164-104">A URL ? https://appsforoffice.microsoft.com/lib/beta/hosted/office.js.</span><span class="sxs-lookup"><span data-stu-id="8a164-104">The URL is https://appsforoffice.microsoft.com/lib/beta/hosted/office.js.</span></span> <span data-ttu-id="8a164-105">Se voc? estiver usando o TypeScript ou se seu editor de c?digo usa um arquivo de defini??o do tipo TypeScript para IntelliSense, use https://appsforoffice.microsoft.com/lib/beta/hosted/office.d.ts.</span><span class="sxs-lookup"><span data-stu-id="8a164-105">If you are using TypeScript or your code editor uses a TypeScript type definition file for intellisense, use https://appsforoffice.microsoft.com/lib/beta/hosted/office.d.ts.</span></span>

<span data-ttu-id="8a164-106">A Biblioteca JavaScript do Excel fornece APIs para permitir que seu suplemento adicione valida??o de dados autom?tica a tabelas, colunas, linhas e outros intervalos em uma pasta de trabalho.</span><span class="sxs-lookup"><span data-stu-id="8a164-106">The Excel JavaScript Library provides APIs to enable your add-in to add automatic data validation to tables, columns, rows, and other ranges in a workbook.</span></span> <span data-ttu-id="8a164-107">Para entender os conceitos e a terminologia de valida??o de dados, consulte os artigos a seguir sobre como os usu?rios adicionam valida??o de dados por meio da interface do usu?rio do Excel:</span><span class="sxs-lookup"><span data-stu-id="8a164-107">To understand the concepts and the terminology of data validation, please see the following articles about how users add data validation through the Excel UI:</span></span>

- [<span data-ttu-id="8a164-108">Aplicar valida??o de dados a c?lulas</span><span class="sxs-lookup"><span data-stu-id="8a164-108">Apply data validation to cells</span></span>](https://support.office.com/en-us/article/Apply-data-validation-to-cells-29FECBCC-D1B9-42C1-9D76-EFF3CE5F7249)
- [<span data-ttu-id="8a164-109">Mais sobre valida??o de dados</span><span class="sxs-lookup"><span data-stu-id="8a164-109">More on data validation</span></span>](https://microsoft.sharepoint.com/:p:/r/teams/oext/_layouts/15/Doc.aspx?sourcedoc=%7B51143964-d52c-429d-bfac-c7495473d536%7D&action=edit)
- [<span data-ttu-id="8a164-110">Descri??o e exemplos de valida??o de dados no Excel</span><span class="sxs-lookup"><span data-stu-id="8a164-110">Description and examples of data validation in Excel</span></span>](https://support.microsoft.com/en-us/help/211485/description-and-examples-of-data-validation-in-excel)

## <a name="programmatic-control-of-data-validation"></a><span data-ttu-id="8a164-111">Controle program?tico de valida??o de dados</span><span class="sxs-lookup"><span data-stu-id="8a164-111">Programmatic control of data validation</span></span>

<span data-ttu-id="8a164-112">A propriedade`Range.dataValidation`, a qual usa um objeto[DataValidation](https://dev.office.com/reference/add-ins/excel/datavalidation), ? o ponto de entrada para o controle program?tico de valida??o de dados no Excel.</span><span class="sxs-lookup"><span data-stu-id="8a164-112">The `Range.dataValidation` property, which takes a [DataValidation](https://dev.office.com/reference/add-ins/excel/datavalidation) object, is the entry point for programmatic control of data validation in Excel.</span></span> <span data-ttu-id="8a164-113">Existem cinco propriedades para o objeto `DataValidation`:</span><span class="sxs-lookup"><span data-stu-id="8a164-113">There are five properties to the `DataValidation` object:</span></span>

- <span data-ttu-id="8a164-114">`rule` ? Define o que constitui dados v?lidos para o intervalo.</span><span class="sxs-lookup"><span data-stu-id="8a164-114">`rule` &#8212; Defines what constitutes valid data for the range.</span></span> <span data-ttu-id="8a164-115">Consulte [DataValidationRule](https://dev.office.com/reference/add-ins/excel/datavalidationrule).</span><span class="sxs-lookup"><span data-stu-id="8a164-115">See [DataValidationRule](https://dev.office.com/reference/add-ins/excel/datavalidationrule).</span></span>
- <span data-ttu-id="8a164-116">`errorAlert` ? Especifica se um erro ser? exibido caso o usu?rio insira dados inv?lidos e define o texto, o t?tulo e o estilo do alerta, por exemplo: **Informativo**, **Aten??o**e **Pare**.</span><span class="sxs-lookup"><span data-stu-id="8a164-116">`errorAlert` &#8212; Specifies whether an error pops up if the user enters invalid data, and defines the alert text, title, and style; for example, **Informational**, **Warning**, and **Stop**.</span></span> <span data-ttu-id="8a164-117">Consulte [DataValidationErrorAlert](https://dev.office.com/reference/add-ins/excel/datavalidationerroralert).</span><span class="sxs-lookup"><span data-stu-id="8a164-117">See [DataValidationErrorAlert](https://dev.office.com/reference/add-ins/excel/datavalidationerroralert).</span></span>
- <span data-ttu-id="8a164-118">`prompt` ? Especifica se uma solicita??o ser? exibida quando o usu?rio passar o mouse sobre o intervalo e define a mensagem da solicita??o.</span><span class="sxs-lookup"><span data-stu-id="8a164-118">`prompt` &#8212; Specifies whether a prompt appears when the user hovers over the range and defines the prompt message.</span></span> <span data-ttu-id="8a164-119">Consulte [DataValidationPrompt](https://dev.office.com/reference/add-ins/excel/datavalidationprompt).</span><span class="sxs-lookup"><span data-stu-id="8a164-119">See [DataValidationPrompt](https://dev.office.com/reference/add-ins/excel/datavalidationprompt).</span></span>
- <span data-ttu-id="8a164-120">`ignoreBlanks` ? Especifica se a regra de valida??o de dados se aplica a c?lulas em branco no intervalo.</span><span class="sxs-lookup"><span data-stu-id="8a164-120">`ignoreBlanks` &#8212; Specifies whether the data validation rule applies to blank cells in the range.</span></span> <span data-ttu-id="8a164-121">Padr?es para `true`.</span><span class="sxs-lookup"><span data-stu-id="8a164-121">Defaults to `true`.</span></span>
- <span data-ttu-id="8a164-122">`type` ? Uma identifica??o somente leitura do tipo de valida??o, como WholeNumber, Date, TextLength, etc. Ela ? definida indiretamente ao se definir a propriedade `rule`.</span><span class="sxs-lookup"><span data-stu-id="8a164-122">`type` &#8212; A read-only identification of the validation type, such as WholeNumber, Date, TextLength, etc. It is set indirectly when you set the `rule` property.</span></span>

> [!NOTE]
> <span data-ttu-id="8a164-123">A valida??o de dados adicionada programaticamente se comporta exatamente como a valida??o de dados adicionada manualmente.</span><span class="sxs-lookup"><span data-stu-id="8a164-123">Data validation added programmatically behaves just like manually added data validation.</span></span> <span data-ttu-id="8a164-124">Em particular, observe que a valida??o de dados s? ? acionada se o usu?rio inserir um valor diretamente em uma c?lula ou copiar e colar uma c?lula de outro local da pasta de trabalho e escolher a op??o de colar**Valores**.</span><span class="sxs-lookup"><span data-stu-id="8a164-124">In particular, note that data validation is triggered only if the user directly enters a value into a cell or copies and pastes a cell from elsewhere in the workbook and chooses the **Values** paste option.</span></span> <span data-ttu-id="8a164-125">Se o usu?rio copiar uma c?lula e executar a a??o de colar sem formata??o em um intervalo com valida??o de dados, a valida??o n?o ser? acionada.</span><span class="sxs-lookup"><span data-stu-id="8a164-125">If the user copies a cell and does a plain paste into a range with data validation, validation is not triggered.</span></span>

### <a name="creating-validation-rules"></a><span data-ttu-id="8a164-126">Criando regras de valida??o</span><span class="sxs-lookup"><span data-stu-id="8a164-126">Creating validation rules</span></span>

<span data-ttu-id="8a164-127">Para adicionar valida??o de dados a um intervalo, seu c?digo deve definir propriedade `rule` do objeto `DataValidation` em `Range.dataValidation`.</span><span class="sxs-lookup"><span data-stu-id="8a164-127">To add data validation to a range, your code must set the `rule` property of the `DataValidation` object in `Range.dataValidation`.</span></span> <span data-ttu-id="8a164-128">Usa-se um objeto [DataValidationRule](https://dev.office.com/reference/add-ins/excel/datavalidationrule) que tem sete propriedades opcionais.</span><span class="sxs-lookup"><span data-stu-id="8a164-128">This takes a [DataValidationRule](https://dev.office.com/reference/add-ins/excel/datavalidationrule) object which has seven optional properties.</span></span> <span data-ttu-id="8a164-129">*N?o pode haver mais do que uma dessas propriedades presente em qualquer objeto `DataValidationRule`.*</span><span class="sxs-lookup"><span data-stu-id="8a164-129">*No more than one of these properties may be present in any `DataValidationRule` object.*</span></span> <span data-ttu-id="8a164-130">A propriedade inclu?da por voc? determina o tipo de valida??o.</span><span class="sxs-lookup"><span data-stu-id="8a164-130">The property that you include determines the type of validation.</span></span>

#### <a name="basic-and-datetime-validation-rule-types"></a><span data-ttu-id="8a164-131">Tipos de regra de valida??o B?sico e DateTime</span><span class="sxs-lookup"><span data-stu-id="8a164-131">Basic and DateTime validation rule types</span></span>

<span data-ttu-id="8a164-132">As tr?s primeiras propriedades `DataValidationRule` (isto ?, tipos de regra de valida??o) usam um objeto [BasicDataValidation](https://dev.office.com/reference/add-ins/excel/basicdatavalidation) como seu valor.</span><span class="sxs-lookup"><span data-stu-id="8a164-132">The first three `DataValidationRule` properties (i.e., validation rule types) take a [BasicDataValidation](https://dev.office.com/reference/add-ins/excel/basicdatavalidation) object as their value.</span></span>

- <span data-ttu-id="8a164-133">`wholeNumber` ? Requer um n?mero inteiro, al?m de qualquer outra valida??o especificada pelo objeto `BasicDataValidation`.</span><span class="sxs-lookup"><span data-stu-id="8a164-133">`wholeNumber` &#8212; Requires a whole number in addition to any other validation specified by the `BasicDataValidation` object.</span></span>
- <span data-ttu-id="8a164-134">`decimal` ? Requer um n?mero decimal, al?m de qualquer outra valida??o especificada pelo objeto `BasicDataValidation`.</span><span class="sxs-lookup"><span data-stu-id="8a164-134">`decimal` &#8212; Requires a decimal number in addition to any other validation specified by the `BasicDataValidation` object.</span></span>
- <span data-ttu-id="8a164-135">`textLength` ? Aplica os detalhes de valida??o no objeto `BasicDataValidation` ao *comprimento* do valor da c?lula.</span><span class="sxs-lookup"><span data-stu-id="8a164-135">`textLength` &#8212; Applies the validation details in the `BasicDataValidation` object to the *length* of the cell's value.</span></span>

<span data-ttu-id="8a164-136">Este ? um exemplo de cria??o de uma regra de valida??o.</span><span class="sxs-lookup"><span data-stu-id="8a164-136">Here is an example of creating a validation rule.</span></span> <span data-ttu-id="8a164-137">Sobre este c?digo, observe o seguinte:</span><span class="sxs-lookup"><span data-stu-id="8a164-137">Note the following about this code:</span></span>

- <span data-ttu-id="8a164-138">O `operator`  ? o operador bin?rio ?GreaterThan?.</span><span class="sxs-lookup"><span data-stu-id="8a164-138">The `operator` is the binary operator "GreaterThan".</span></span> <span data-ttu-id="8a164-139">Sempre for usado um operador bin?rio, o valor que o usu?rio tentar inserir na c?lula ? o operando esquerdo, e o valor especificado em `formula1` ? o operando direito.</span><span class="sxs-lookup"><span data-stu-id="8a164-139">Whenever you use a binary operator, the value that the user tries to enter in the cell is the left-hand operand and the value specified in `formula1` is the right-hand operand.</span></span> <span data-ttu-id="8a164-140">Portanto, essa regra diz que apenas n?meros inteiros maiores que 0 s?o v?lidos.</span><span class="sxs-lookup"><span data-stu-id="8a164-140">So this rule says that only whole numbers that are greater than 0 are valid.</span></span> 
- <span data-ttu-id="8a164-141">O `formula1` ? um n?mero embutido em c?digo.</span><span class="sxs-lookup"><span data-stu-id="8a164-141">The `formula1` is a hard-coded number.</span></span> <span data-ttu-id="8a164-142">No momento da codifica??o, caso n?o saiba qual deve ser o valor, voc? tamb?m poder? usar uma f?rmula do Excel (como uma sequ?ncia de caracteres) para o valor.</span><span class="sxs-lookup"><span data-stu-id="8a164-142">If you don't know at coding time what the value should be, you can also use an Excel formula (as a string) for the value.</span></span> <span data-ttu-id="8a164-143">Por exemplo, ?= A3? e ?= SUM(A4, B5)? tamb?m podem ser valores de `formula1`.</span><span class="sxs-lookup"><span data-stu-id="8a164-143">For example, "=A3" and "=SUM(A4,B5)" could also be values of `formula1`.</span></span>

```js
Excel.run(function (context) {
    var sheet = context.workbook.worksheets.getActiveWorksheet();
    var range = sheet.getRange("B2:C5");
   
    range.dataValidation.rule = {
            wholeNumber: {
                formula1: 0,
                operator: "GreaterThan"
            }
        };

    return context.sync();
})
```

<span data-ttu-id="8a164-144">Consulte [BasicDataValidation](https://dev.office.com/reference/add-ins/excel/basicdatavalidation) para obter uma lista dos outros operadores bin?rios.</span><span class="sxs-lookup"><span data-stu-id="8a164-144">See [BasicDataValidation](https://dev.office.com/reference/add-ins/excel/basicdatavalidation) for a list of the other binary operators.</span></span> 

<span data-ttu-id="8a164-145">Existem tamb?m dois operadores tern?rios: ?Between? e ?NotBetween?.</span><span class="sxs-lookup"><span data-stu-id="8a164-145">There are also two ternary operators: "Between" and "NotBetween".</span></span> <span data-ttu-id="8a164-146">Para us?-los, voc? deve especificar a propriedade opcional `formula2`.</span><span class="sxs-lookup"><span data-stu-id="8a164-146">To use these, you must specify the optional `formula2` property.</span></span> <span data-ttu-id="8a164-147">Os valores `formula1` e `formula2` s?o os operandos delimitadores.</span><span class="sxs-lookup"><span data-stu-id="8a164-147">The `formula1` and `formula2` values are the bounding operands.</span></span> <span data-ttu-id="8a164-148">O valor que o usu?rio tentar inserir na c?lula ? o terceiro operando (avaliado).</span><span class="sxs-lookup"><span data-stu-id="8a164-148">The value that the user tries to enter in the cell is the third (evaluated) operand.</span></span> <span data-ttu-id="8a164-149">A seguir, h? um exemplo de uso do operador ?Between?:</span><span class="sxs-lookup"><span data-stu-id="8a164-149">The following is an example of using the "Between" operator:</span></span>

```js
Excel.run(function (context) {
    var sheet = context.workbook.worksheets.getActiveWorksheet();
    var range = sheet.getRange("B2:C5");
   
    range.dataValidation.rule = {
            decimal: {
                formula1: 0,
                formula2: 100,
                operator: "Between"
            }
        };

    return context.sync();
})
```

<span data-ttu-id="8a164-150">As pr?ximas duas propriedades da regra usam o objeto [DateTimeDataValidation](https://dev.office.com/reference/add-ins/excel/basicdatavalidation) como seu valor.</span><span class="sxs-lookup"><span data-stu-id="8a164-150">The next two rule properties take a [DateTimeDataValidation](https://dev.office.com/reference/add-ins/excel/basicdatavalidation) object as their value.</span></span>

- `date`
- `time`

<span data-ttu-id="8a164-151">O objeto `DateTimeDataValidation` ? estruturado de forma semelhante ao `BasicDataValidation`: tem as propriedades `formula1`, `formula2`e `operator` e ? usado da mesma maneira.</span><span class="sxs-lookup"><span data-stu-id="8a164-151">The `DateTimeDataValidation` object is structured similarly to the `BasicDataValidation`: it has the properties `formula1`, `formula2`, and `operator`, and is used in the same way.</span></span> <span data-ttu-id="8a164-152">A diferen?a ? que voc? n?o pode usar um n?mero nas propriedades da f?rmula, mas pode inserir uma sequ?ncia de caracteres [ISO 8606 datetime](https://www.iso.org/iso-8601-date-and-time-format.html) (ou uma f?rmula do Excel).</span><span class="sxs-lookup"><span data-stu-id="8a164-152">The difference is that you cannot use a number in the formula properties, but you can enter a [ISO 8606 datetime](https://www.iso.org/iso-8601-date-and-time-format.html) string (or an Excel formula).</span></span> <span data-ttu-id="8a164-153">A seguir, h? um exemplo que define valores v?lidos como datas na primeira semana de abril de 2018.</span><span class="sxs-lookup"><span data-stu-id="8a164-153">The following is an example that defines valid values as dates in the first week of April, 2018.</span></span> 

```js
Excel.run(function (context) {
    var sheet = context.workbook.worksheets.getActiveWorksheet();
    var range = sheet.getRange("B2:C5");
   
    range.dataValidation.rule = {
            date: {
                formula1: "2018-04-01",
                formula2: "2018-04-08",
                operator: "Between"
            }
        };

    return context.sync();
})
```

#### <a name="list-validation-rule-type"></a><span data-ttu-id="8a164-154">Tipo de regra de valida??o de lista</span><span class="sxs-lookup"><span data-stu-id="8a164-154">List validation rule type</span></span>

<span data-ttu-id="8a164-155">Use a propriedade `list` no objeto `DataValidationRule` para especificar que os ?nicos valores v?lidos sejam aqueles de uma lista finita.</span><span class="sxs-lookup"><span data-stu-id="8a164-155">Use the `list` property in the `DataValidationRule` object to specify that the only valid values are those from a finite list.</span></span> <span data-ttu-id="8a164-156">H? um exemplo a seguir.</span><span class="sxs-lookup"><span data-stu-id="8a164-156">The following is an example.</span></span> <span data-ttu-id="8a164-157">Sobre este c?digo, observe o seguinte:</span><span class="sxs-lookup"><span data-stu-id="8a164-157">Note the following about this code:</span></span>

- <span data-ttu-id="8a164-158">Ele pressup?e que h? uma planilha chamada ?Nomes? e que os valores no intervalo ?A1: A3? s?o nomes.</span><span class="sxs-lookup"><span data-stu-id="8a164-158">It assumes that there is a worksheet named "Names" and that the values in the range "A1:A3" are names.</span></span>
- <span data-ttu-id="8a164-159">A propriedade `source` especifica a lista de valores v?lidos.</span><span class="sxs-lookup"><span data-stu-id="8a164-159">The `source` property specifies the list of valid values.</span></span> <span data-ttu-id="8a164-160">O intervalo com os nomes foi atribu?do a ela.</span><span class="sxs-lookup"><span data-stu-id="8a164-160">The range with the names has been assigned to it.</span></span> <span data-ttu-id="8a164-161">Tamb?m ? poss?vel atribuir uma lista delimitada por v?rgula, por exemplo: ?Sue, Ricky, Liz?.</span><span class="sxs-lookup"><span data-stu-id="8a164-161">You can also assign a comma-delimited list; for example: "Sue, Ricky, Liz".</span></span> 
- <span data-ttu-id="8a164-162">A propriedade `inCellDropDown` especifica se um controle suspenso aparecer? na c?lula quando o usu?rio selecion?-lo.</span><span class="sxs-lookup"><span data-stu-id="8a164-162">The `inCellDropDown` property specifies whether a drop-down control will appear in the cell when the user selects it.</span></span> <span data-ttu-id="8a164-163">Se definida como `true`, a lista suspensa aparece com a lista de valores de `source`.</span><span class="sxs-lookup"><span data-stu-id="8a164-163">If set to `true`, then the drop-down appears with the list of values from the `source`.</span></span>

```js
Excel.run(function (context) {
    var sheet = context.workbook.worksheets.getActiveWorksheet();
    var range = sheet.getRange("B2:C5");   
    var nameSourceRange = context.workbook.worksheets.getItem("Names").getRange("A1:A3");

    range.dataValidation.rule = {
        list: {
            inCellDropDown: true,
            source: nameSourceRange
        }
    };

    return context.sync();
})
```

#### <a name="custom-validation-rule-type"></a><span data-ttu-id="8a164-164">Tipo de regra de valida??o personalizada</span><span class="sxs-lookup"><span data-stu-id="8a164-164">Custom validation rule type</span></span>

<span data-ttu-id="8a164-165">Use a propriedade `custom` no objeto `DataValidationRule` para especificar uma f?rmula de valida??o personalizada.</span><span class="sxs-lookup"><span data-stu-id="8a164-165">Use the `custom` property in the `DataValidationRule` object to specify a custom validation formula.</span></span> <span data-ttu-id="8a164-166">H? um exemplo a seguir.</span><span class="sxs-lookup"><span data-stu-id="8a164-166">The following is an example.</span></span> <span data-ttu-id="8a164-167">Sobre este c?digo, observe o seguinte:</span><span class="sxs-lookup"><span data-stu-id="8a164-167">Note the following about this code:</span></span>

- <span data-ttu-id="8a164-168">Ele pressup?e que h? uma tabela de duas colunas com colunas **Nome do Atleta** e **Coment?rios** nas colunas A e B da planilha.</span><span class="sxs-lookup"><span data-stu-id="8a164-168">It assumes there is a two-column table with columns **Athlete Name** and **Comments** in the A and B columns of the worksheet.</span></span>
- <span data-ttu-id="8a164-169">Para reduzir a verbosidade na coluna **Coment?rios**, ele faz com que os dados que incluem o nome do atleta se tornem inv?lidos.</span><span class="sxs-lookup"><span data-stu-id="8a164-169">To reduce verbosity in the **Comments** column, it makes data that includes the athlete's name invalid.</span></span>
- <span data-ttu-id="8a164-170">`SEARCH(A2,B2)` retorna a posi??o inicial, na sequ?ncia de caracteres B2, da sequ?ncia de caracteres em A2.</span><span class="sxs-lookup"><span data-stu-id="8a164-170">`SEARCH(A2,B2)` returns the starting position, in string in B2, of the string in A2.</span></span> <span data-ttu-id="8a164-171">Se A2 n?o estiver contida em B2, ele n?o retornar? um n?mero.</span><span class="sxs-lookup"><span data-stu-id="8a164-171">If A2 is not contained in B2, it does not return a number.</span></span> <span data-ttu-id="8a164-172">`ISNUMBER()` retorna um booleano.</span><span class="sxs-lookup"><span data-stu-id="8a164-172">Returns a `ISNUMBER()`.</span></span> <span data-ttu-id="8a164-173">Ent?o a propriedade `formula` diz que dados v?lidos da coluna **Coment?rios** s?o dados que n?o incluem a sequ?ncia de caracteres na coluna **Nome do Atleta**.</span><span class="sxs-lookup"><span data-stu-id="8a164-173">So the `formula` property says that valid data for the **Comment** column is data that does not include the string in the **Athlete Name** column.</span></span>

```js
Excel.run(function (context) {
    var sheet = context.workbook.worksheets.getActiveWorksheet();
    var range = sheet.getRange("B2:C5");   
    var commentsRange = sheet.tables.getItem("AthletesTable").columns.getItem("Comments").getDataBodyRange();

    commentsRange.dataValidation.rule = {
            custom: {
                formula: "=NOT(ISNUMBER(SEARCH(A2,B2)))"
            }
        };

    return context.sync();
})
```

### <a name="create-validation-error-alerts"></a><span data-ttu-id="8a164-174">Criar alertas de erro de valida??o</span><span class="sxs-lookup"><span data-stu-id="8a164-174">Create validation error alerts</span></span>

<span data-ttu-id="8a164-175">? poss?vel criar um alerta de erro personalizado que aparecer? quando um usu?rio tentar inserir dados inv?lidos em uma c?lula.</span><span class="sxs-lookup"><span data-stu-id="8a164-175">You can a create custom error alert that appears when a user tries to enter invalid data in a cell.</span></span> <span data-ttu-id="8a164-176">H? um exemplo simples a seguir.</span><span class="sxs-lookup"><span data-stu-id="8a164-176">The following is a simple example:</span></span> <span data-ttu-id="8a164-177">Sobre este c?digo, observe o seguinte:</span><span class="sxs-lookup"><span data-stu-id="8a164-177">Note the following about this code:</span></span>

- <span data-ttu-id="8a164-178">A propriedade `style` determina se o usu?rio recebe um alerta informativo, um aviso ou um alerta do tipo ?pare?.</span><span class="sxs-lookup"><span data-stu-id="8a164-178">The `style` property determines whether the user gets an informational alert, a warning, or a "stop" alert.</span></span> <span data-ttu-id="8a164-179">Somente `Stop` impede de verdade que o usu?rio adicione dados inv?lidos.</span><span class="sxs-lookup"><span data-stu-id="8a164-179">Only `Stop` actually prevents the user from adding invalid data.</span></span> <span data-ttu-id="8a164-180">O pop-up para `Warning` e `Information` tem op??es que permitem que o usu?rio insira os dados inv?lidos.</span><span class="sxs-lookup"><span data-stu-id="8a164-180">The pop-up for `Warning` and `Information` has options that allow the user enter the invalid data anyway.</span></span>
- <span data-ttu-id="8a164-181">A propriedade `showAlert` se torna padr?o para `true`.</span><span class="sxs-lookup"><span data-stu-id="8a164-181">The `showAlert` property defaults to `true`.</span></span> <span data-ttu-id="8a164-182">Isso significa que o host do Excel exibir? um alerta pop-up gen?rico (do tipo `Stop`) a menos que seja criado um alerta personalizado que defina `showAlert` para `false` ou defina uma mensagem, um t?tulo e um estilo personalizados.</span><span class="sxs-lookup"><span data-stu-id="8a164-182">This means that the Excel host will pop-up a generic alert (of type `Stop`) unless you create a custom alert which either sets `showAlert` to `false` or sets a custom message, title, and style.</span></span> <span data-ttu-id="8a164-183">Esse c?digo define uma mensagem personalizada e um t?tulo.</span><span class="sxs-lookup"><span data-stu-id="8a164-183">This code sets a custom message and title.</span></span>


```js
Excel.run(function (context) {
    var sheet = context.workbook.worksheets.getActiveWorksheet();
    var range = sheet.getRange("B2:C5");
   
    range.dataValidation.errorAlert = {
            message: "Sorry, only positive whole numbers are allowed",
            showAlert: true, // default is 'true'
            style: "Stop", // other possible values: Warning, Information
            title: "Negative or Decimal Number Entered"
        };
    
    // Set range.dataValidation.rule and optionally .prompt here.

    return context.sync();
})
```

<span data-ttu-id="8a164-184">Para obter mais informa??es, consulte [DataValidationErrorAlert](https://dev.office.com/reference/add-ins/excel/datavalidationerroralert).</span><span class="sxs-lookup"><span data-stu-id="8a164-184">For more information, see [NextRecordset](https://dev.office.com/reference/add-ins/excel/datavalidationerroralert).</span></span>

### <a name="create-validation-prompts"></a><span data-ttu-id="8a164-185">Criar solicita??es de valida??o</span><span class="sxs-lookup"><span data-stu-id="8a164-185">Create validation prompts</span></span>

<span data-ttu-id="8a164-186">? poss?vel criar uma solicita??o de instru??o que aparece quando um usu?rio seleciona ou passa o mouse sobre uma c?lula na qual a valida??o de dados foi aplicada.</span><span class="sxs-lookup"><span data-stu-id="8a164-186">You can create an instructional prompt that appears when a user hovers over, or selects, a cell to which data validation has been applied.</span></span> <span data-ttu-id="8a164-187">H? um exemplo a seguir:</span><span class="sxs-lookup"><span data-stu-id="8a164-187">The following is an example:</span></span>

```js
Excel.run(function (context) {
    var sheet = context.workbook.worksheets.getActiveWorksheet();
    var range = sheet.getRange("B2:C5");
   
    range.dataValidation.prompt = {
            message: "Please enter a positive whole number.",
            showPrompt: true, // default is 'false'
            title: "Positive Whole Numbers Only."
        };
    
    // Set range.dataValidation.rule and optionally .errorAlert here.

    return context.sync();
})
```

<span data-ttu-id="8a164-188">Para obter mais informa??es, consulte [DataValidationPrompt](https://dev.office.com/reference/add-ins/excel/datavalidationprompt).</span><span class="sxs-lookup"><span data-stu-id="8a164-188">For more information, see [NextRecordset](https://dev.office.com/reference/add-ins/excel/datavalidationprompt).</span></span>

### <a name="remove-data-validation-from-a-range"></a><span data-ttu-id="8a164-189">Remover a valida??o de dados de um intervalo</span><span class="sxs-lookup"><span data-stu-id="8a164-189">Remove data validation from a range</span></span>

<span data-ttu-id="8a164-190">Para remover a valida??o de dados de um intervalo, chame o m?todo [Range.dataValidation.clear()](https://dev.office.com/reference/add-ins/excel/datavalidation#clear).</span><span class="sxs-lookup"><span data-stu-id="8a164-190">To remove data validation from a range, call the  [Range.dataValidation.clear()](https://dev.office.com/reference/add-ins/excel/datavalidation#clear) method.</span></span>

```js
myrange.dataValidation.clear()
```

<span data-ttu-id="8a164-191">N?o ? necess?rio que o intervalo limpo seja exatamente o mesmo de um intervalo no qual a valida??o de dados foi adicionada.</span><span class="sxs-lookup"><span data-stu-id="8a164-191">It isn't necessary that the range you clear is exactly the same range as a range on which you added data validation.</span></span> <span data-ttu-id="8a164-192">Se n?o for, apenas as c?lulas sobrepostas, se houver, dos dois intervalos ser?o limpas.</span><span class="sxs-lookup"><span data-stu-id="8a164-192">If it isn't, only the overlapping cells, if any, of the two ranges are cleared.</span></span> 

> [!NOTE]
> <span data-ttu-id="8a164-193">A limpeza da valida??o de dados de um intervalo tamb?m limpar? qualquer valida??o de dados que um usu?rio tenha adicionado manualmente ao intervalo.</span><span class="sxs-lookup"><span data-stu-id="8a164-193">Clearing data validation from a range will also clear any data validation that a user has added manually to the range.</span></span>

## <a name="see-also"></a><span data-ttu-id="8a164-194">Confira tamb?m</span><span class="sxs-lookup"><span data-stu-id="8a164-194">See also</span></span>

- [<span data-ttu-id="8a164-195">Principais conceitos da API JavaScript do Excel</span><span class="sxs-lookup"><span data-stu-id="8a164-195">Excel JavaScript API core concepts</span></span>](excel-add-ins-core-concepts.md)
- [<span data-ttu-id="8a164-196">Objeto DataValidation (API JavaScript para Excel)</span><span class="sxs-lookup"><span data-stu-id="8a164-196">Worksheet Object (JavaScript API for Excel)</span></span>](https://dev.office.com/reference/add-ins/excel/datavalidation)
- [<span data-ttu-id="8a164-197">Objeto Range (API JavaScript para Excel)</span><span class="sxs-lookup"><span data-stu-id="8a164-197">Range Object (JavaScript API for Excel)</span></span>](https://dev.office.com/reference/add-ins/excel/range)



 
