---
title: Adicionar validação de dados a intervalos do Excel
description: ''
ms.date: 04/13/2018
ms.openlocfilehash: af965df4a1aece5b7f8d5ea89664519b576a4850
ms.sourcegitcommit: 4de2a1b62ccaa8e51982e95537fc9f52c0c5e687
ms.translationtype: HT
ms.contentlocale: pt-BR
ms.lasthandoff: 08/10/2018
ms.locfileid: "22925308"
---
# <a name="add-data-validation-to-excel-ranges-preview"></a><span data-ttu-id="647b0-102">Adicionar validação de dados a intervalos do Excel (versão prévia)</span><span class="sxs-lookup"><span data-stu-id="647b0-102">Add data validation to Excel ranges (Preview)</span></span>

> [!NOTE]
> <span data-ttu-id="647b0-103">Enquanto as APIs de validação de dados estão em versão prévia, você deve carregar a versão beta da biblioteca JavaScript do Office para usá-las.</span><span class="sxs-lookup"><span data-stu-id="647b0-103">While the data validation APIs are in preview, you must load the beta version of the Office JavaScript library to use them.</span></span> <span data-ttu-id="647b0-104">A URL é https://appsforoffice.microsoft.com/lib/beta/hosted/office.js.</span><span class="sxs-lookup"><span data-stu-id="647b0-104">The URL is https://appsforoffice.microsoft.com/lib/beta/hosted/office.js.</span></span> <span data-ttu-id="647b0-105">Se você estiver usando o TypeScript ou se seu editor de código usa um arquivo de definição do tipo TypeScript para IntelliSense, use https://appsforoffice.microsoft.com/lib/beta/hosted/office.d.ts.</span><span class="sxs-lookup"><span data-stu-id="647b0-105">If you are using TypeScript or your code editor uses a TypeScript type definition file for intellisense, use https://appsforoffice.microsoft.com/lib/beta/hosted/office.d.ts.</span></span>

> [!NOTE]
> <span data-ttu-id="647b0-106">Embora as APIs de validação de dados estejam em versão prévia, os links neste artigo para a referência da API não funcionarão.</span><span class="sxs-lookup"><span data-stu-id="647b0-106">While the data validation APIs are in preview, the links in this article to API reference will not work.</span></span> <span data-ttu-id="647b0-107">Enquanto isso, você pode usar a [referência da API do Excel de rascunho](https://github.com/OfficeDev/office-js-docs/tree/ExcelJs_OpenSpec/reference/excel).</span><span class="sxs-lookup"><span data-stu-id="647b0-107">In the meantime, you can use the [draft Excel API reference](https://github.com/OfficeDev/office-js-docs/tree/ExcelJs_OpenSpec/reference/excel).</span></span>

<span data-ttu-id="647b0-108">A Biblioteca JavaScript do Excel fornece APIs para permitir que seu suplemento adicione validação de dados automática a tabelas, colunas, linhas e outros intervalos em uma pasta de trabalho.</span><span class="sxs-lookup"><span data-stu-id="647b0-108">The Excel JavaScript Library provides APIs to enable your add-in to add automatic data validation to tables, columns, rows, and other ranges in a workbook.</span></span> <span data-ttu-id="647b0-109">Para entender os conceitos e a terminologia de validação de dados, consulte os artigos a seguir sobre como os usuários adicionam validação de dados por meio da interface do usuário do Excel:</span><span class="sxs-lookup"><span data-stu-id="647b0-109">To understand the concepts and the terminology of data validation, please see the following articles about how users add data validation through the Excel UI:</span></span>

- [<span data-ttu-id="647b0-110">Aplicar validação de dados a células</span><span class="sxs-lookup"><span data-stu-id="647b0-110">Apply data validation to cells</span></span>](https://support.office.com/article/Apply-data-validation-to-cells-29FECBCC-D1B9-42C1-9D76-EFF3CE5F7249)
- [<span data-ttu-id="647b0-111">Mais sobre validação de dados</span><span class="sxs-lookup"><span data-stu-id="647b0-111">More on data validation</span></span>](https://support.office.com/article/More-on-data-validation-f38dee73-9900-4ca6-9301-8a5f6e1f0c4c)
- [<span data-ttu-id="647b0-112">Descrição e exemplos de validação de dados no Excel</span><span class="sxs-lookup"><span data-stu-id="647b0-112">Description and examples of data validation in Excel</span></span>](https://support.microsoft.com/help/211485/description-and-examples-of-data-validation-in-excel)

## <a name="programmatic-control-of-data-validation"></a><span data-ttu-id="647b0-113">Controle programático de validação de dados</span><span class="sxs-lookup"><span data-stu-id="647b0-113">Programmatic control of data validation</span></span>

<span data-ttu-id="647b0-114">A propriedade`Range.dataValidation`, a qual usa um objeto[DataValidation](https://docs.microsoft.com/javascript/api/excel/excel.datavalidation), é o ponto de entrada para o controle programático de validação de dados no Excel.</span><span class="sxs-lookup"><span data-stu-id="647b0-114">The `Range.dataValidation` property, which takes a [DataValidation](https://docs.microsoft.com/javascript/api/excel/excel.datavalidation) object, is the entry point for programmatic control of data validation in Excel.</span></span> <span data-ttu-id="647b0-115">Existem cinco propriedades para o objeto `DataValidation`:</span><span class="sxs-lookup"><span data-stu-id="647b0-115">There are five properties to the `DataValidation` object:</span></span>

- <span data-ttu-id="647b0-116">`rule` – Define o que constitui dados válidos para o intervalo.</span><span class="sxs-lookup"><span data-stu-id="647b0-116">`rule` &#8212; Defines what constitutes valid data for the range.</span></span> <span data-ttu-id="647b0-117">Consulte [DataValidationRule](https://docs.microsoft.com/javascript/api/excel/excel.datavalidationrule).</span><span class="sxs-lookup"><span data-stu-id="647b0-117">See [DataValidationRule](https://docs.microsoft.com/javascript/api/excel/excel.datavalidationrule).</span></span>
- <span data-ttu-id="647b0-118">`errorAlert` – Especifica se um erro será exibido caso o usuário insira dados inválidos e define o texto, o título e o estilo do alerta, por exemplo: **Informativo**, **Atenção**e **Pare**.</span><span class="sxs-lookup"><span data-stu-id="647b0-118">`errorAlert` &#8212; Specifies whether an error pops up if the user enters invalid data, and defines the alert text, title, and style; for example, **Informational**, **Warning**, and **Stop**.</span></span> <span data-ttu-id="647b0-119">Consulte [DataValidationErrorAlert](https://docs.microsoft.com/javascript/api/excel/excel.datavalidationerroralert).</span><span class="sxs-lookup"><span data-stu-id="647b0-119">See [DataValidationErrorAlert](https://docs.microsoft.com/javascript/api/excel/excel.datavalidationerroralert).</span></span>
- <span data-ttu-id="647b0-120">`prompt` – Especifica se uma solicitação será exibida quando o usuário passar o mouse sobre o intervalo e define a mensagem da solicitação.</span><span class="sxs-lookup"><span data-stu-id="647b0-120">`prompt` &#8212; Specifies whether a prompt appears when the user hovers over the range and defines the prompt message.</span></span> <span data-ttu-id="647b0-121">Consulte [DataValidationPrompt](https://docs.microsoft.com/javascript/api/excel/excel.datavalidationprompt).</span><span class="sxs-lookup"><span data-stu-id="647b0-121">See [DataValidationPrompt](https://docs.microsoft.com/javascript/api/excel/excel.datavalidationprompt).</span></span>
- <span data-ttu-id="647b0-122">`ignoreBlanks` – Especifica se a regra de validação de dados se aplica a células em branco no intervalo.</span><span class="sxs-lookup"><span data-stu-id="647b0-122">`ignoreBlanks` &#8212; Specifies whether the data validation rule applies to blank cells in the range.</span></span> <span data-ttu-id="647b0-123">Padrões para `true`.</span><span class="sxs-lookup"><span data-stu-id="647b0-123">Defaults to `true`.</span></span>
- <span data-ttu-id="647b0-124">`type` – Uma identificação somente leitura do tipo de validação, como WholeNumber, Date, TextLength, etc. Ela é definida indiretamente ao se definir a propriedade `rule`.</span><span class="sxs-lookup"><span data-stu-id="647b0-124">`type` &#8212; A read-only identification of the validation type, such as WholeNumber, Date, TextLength, etc. It is set indirectly when you set the `rule` property.</span></span>

> [!NOTE]
> <span data-ttu-id="647b0-125">A validação de dados adicionada programaticamente se comporta exatamente como a validação de dados adicionada manualmente.</span><span class="sxs-lookup"><span data-stu-id="647b0-125">Data validation added programmatically behaves just like manually added data validation.</span></span> <span data-ttu-id="647b0-126">Em particular, observe que a validação de dados só é acionada se o usuário inserir um valor diretamente em uma célula ou copiar e colar uma célula de outro local da pasta de trabalho e escolher a opção de colar**Valores**.</span><span class="sxs-lookup"><span data-stu-id="647b0-126">In particular, note that data validation is triggered only if the user directly enters a value into a cell or copies and pastes a cell from elsewhere in the workbook and chooses the **Values** paste option.</span></span> <span data-ttu-id="647b0-127">Se o usuário copiar uma célula e executar a ação de colar sem formatação em um intervalo com validação de dados, a validação não será acionada.</span><span class="sxs-lookup"><span data-stu-id="647b0-127">If the user copies a cell and does a plain paste into a range with data validation, validation is not triggered.</span></span>

### <a name="creating-validation-rules"></a><span data-ttu-id="647b0-128">Criando regras de validação</span><span class="sxs-lookup"><span data-stu-id="647b0-128">Creating validation rules</span></span>

<span data-ttu-id="647b0-129">Para adicionar validação de dados a um intervalo, seu código deve definir propriedade `rule` do objeto `DataValidation` em `Range.dataValidation`.</span><span class="sxs-lookup"><span data-stu-id="647b0-129">To add data validation to a range, your code must set the `rule` property of the `DataValidation` object in `Range.dataValidation`.</span></span> <span data-ttu-id="647b0-130">Usa-se um objeto [DataValidationRule](https://dev.office.com/reference/add-ins/excel/datavalidationrule) que tem sete propriedades opcionais.</span><span class="sxs-lookup"><span data-stu-id="647b0-130">This takes a [DataValidationRule](https://dev.office.com/reference/add-ins/excel/datavalidationrule) object which has seven optional properties.</span></span> <span data-ttu-id="647b0-131">*Não pode haver mais do que uma dessas propriedades presente em qualquer objeto `DataValidationRule`.*</span><span class="sxs-lookup"><span data-stu-id="647b0-131">*No more than one of these properties may be present in any `DataValidationRule` object.*</span></span> <span data-ttu-id="647b0-132">A propriedade incluída por você determina o tipo de validação.</span><span class="sxs-lookup"><span data-stu-id="647b0-132">The property that you include determines the type of validation.</span></span>

#### <a name="basic-and-datetime-validation-rule-types"></a><span data-ttu-id="647b0-133">Tipos de regra de validação Básico e DateTime</span><span class="sxs-lookup"><span data-stu-id="647b0-133">Basic and DateTime validation rule types</span></span>

<span data-ttu-id="647b0-134">As três primeiras propriedades `DataValidationRule` (isto é, tipos de regra de validação) usam um objeto [BasicDataValidation](https://docs.microsoft.com/javascript/api/excel/excel.basicdatavalidation) como seu valor.</span><span class="sxs-lookup"><span data-stu-id="647b0-134">The first three `DataValidationRule` properties (i.e., validation rule types) take a [BasicDataValidation](https://docs.microsoft.com/javascript/api/excel/excel.basicdatavalidation) object as their value.</span></span>

- <span data-ttu-id="647b0-135">`wholeNumber` – Requer um número inteiro, além de qualquer outra validação especificada pelo objeto `BasicDataValidation`.</span><span class="sxs-lookup"><span data-stu-id="647b0-135">`wholeNumber` &#8212; Requires a whole number in addition to any other validation specified by the `BasicDataValidation` object.</span></span>
- <span data-ttu-id="647b0-136">`decimal` – Requer um número decimal, além de qualquer outra validação especificada pelo objeto `BasicDataValidation`.</span><span class="sxs-lookup"><span data-stu-id="647b0-136">`decimal` &#8212; Requires a decimal number in addition to any other validation specified by the `BasicDataValidation` object.</span></span>
- <span data-ttu-id="647b0-137">`textLength` – Aplica os detalhes de validação no objeto `BasicDataValidation` ao *comprimento* do valor da célula.</span><span class="sxs-lookup"><span data-stu-id="647b0-137">`textLength` &#8212; Applies the validation details in the `BasicDataValidation` object to the *length* of the cell's value.</span></span>

<span data-ttu-id="647b0-138">Este é um exemplo de criação de uma regra de validação.</span><span class="sxs-lookup"><span data-stu-id="647b0-138">Here is an example of creating a validation rule.</span></span> <span data-ttu-id="647b0-139">Observe o seguinte sobre este código:</span><span class="sxs-lookup"><span data-stu-id="647b0-139">Note the following about this code:</span></span>

- <span data-ttu-id="647b0-140">O `operator`  é o operador binário “GreaterThan”.</span><span class="sxs-lookup"><span data-stu-id="647b0-140">The `operator` is the binary operator "GreaterThan".</span></span> <span data-ttu-id="647b0-141">Sempre for usado um operador binário, o valor que o usuário tentar inserir na célula é o operando esquerdo, e o valor especificado em `formula1` é o operando direito.</span><span class="sxs-lookup"><span data-stu-id="647b0-141">Whenever you use a binary operator, the value that the user tries to enter in the cell is the left-hand operand and the value specified in `formula1` is the right-hand operand.</span></span> <span data-ttu-id="647b0-142">Portanto, essa regra diz que apenas números inteiros maiores que 0 são válidos.</span><span class="sxs-lookup"><span data-stu-id="647b0-142">So this rule says that only whole numbers that are greater than 0 are valid.</span></span> 
- <span data-ttu-id="647b0-143">O `formula1` é um número embutido em código.</span><span class="sxs-lookup"><span data-stu-id="647b0-143">The `formula1` is a hard-coded number.</span></span> <span data-ttu-id="647b0-144">No momento da codificação, caso não saiba qual deve ser o valor, você também poderá usar uma fórmula do Excel (como uma sequência de caracteres) para o valor.</span><span class="sxs-lookup"><span data-stu-id="647b0-144">If you don't know at coding time what the value should be, you can also use an Excel formula (as a string) for the value.</span></span> <span data-ttu-id="647b0-145">Por exemplo, “= A3” e “= SUM(A4, B5)” também podem ser valores de `formula1`.</span><span class="sxs-lookup"><span data-stu-id="647b0-145">For example, "=A3" and "=SUM(A4,B5)" could also be values of `formula1`.</span></span>

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

<span data-ttu-id="647b0-146">Consulte [BasicDataValidation](https://docs.microsoft.com/javascript/api/excel/excel.basicdatavalidation) para obter uma lista dos outros operadores binários.</span><span class="sxs-lookup"><span data-stu-id="647b0-146">See [BasicDataValidation](https://docs.microsoft.com/javascript/api/excel/excel.basicdatavalidation) for a list of the other binary operators.</span></span> 

<span data-ttu-id="647b0-147">Existem também dois operadores ternários: “Between” e “NotBetween”.</span><span class="sxs-lookup"><span data-stu-id="647b0-147">There are also two ternary operators: "Between" and "NotBetween".</span></span> <span data-ttu-id="647b0-148">Para usá-los, você deve especificar a propriedade opcional `formula2`.</span><span class="sxs-lookup"><span data-stu-id="647b0-148">To use these, you must specify the optional `formula2` property.</span></span> <span data-ttu-id="647b0-149">Os valores `formula1` e `formula2` são os operandos delimitadores.</span><span class="sxs-lookup"><span data-stu-id="647b0-149">The `formula1` and `formula2` values are the bounding operands.</span></span> <span data-ttu-id="647b0-150">O valor que o usuário tentar inserir na célula é o terceiro operando (avaliado).</span><span class="sxs-lookup"><span data-stu-id="647b0-150">The value that the user tries to enter in the cell is the third (evaluated) operand.</span></span> <span data-ttu-id="647b0-151">A seguir, há um exemplo de uso do operador “Between”:</span><span class="sxs-lookup"><span data-stu-id="647b0-151">The following is an example of using the "Between" operator:</span></span>

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

<span data-ttu-id="647b0-152">As próximas duas propriedades da regra usam o objeto [DateTimeDataValidation](https://docs.microsoft.com/javascript/api/excel/excel.datetimedatavalidation) como seu valor.</span><span class="sxs-lookup"><span data-stu-id="647b0-152">The next two rule properties take a [DateTimeDataValidation](https://docs.microsoft.com/javascript/api/excel/excel.datetimedatavalidation) object as their value.</span></span>

- `date`
- `time`

<span data-ttu-id="647b0-153">O objeto `DateTimeDataValidation` é estruturado de forma semelhante ao `BasicDataValidation`: tem as propriedades `formula1`, `formula2`e `operator` e é usado da mesma maneira.</span><span class="sxs-lookup"><span data-stu-id="647b0-153">The `DateTimeDataValidation` object is structured similarly to the `BasicDataValidation`: it has the properties `formula1`, `formula2`, and `operator`, and is used in the same way.</span></span> <span data-ttu-id="647b0-154">A diferença é que você não pode usar um número nas propriedades da fórmula, mas pode inserir uma sequência de caracteres [ISO 8606 datetime](https://www.iso.org/iso-8601-date-and-time-format.html) (ou uma fórmula do Excel).</span><span class="sxs-lookup"><span data-stu-id="647b0-154">The difference is that you cannot use a number in the formula properties, but you can enter a [ISO 8606 datetime](https://www.iso.org/iso-8601-date-and-time-format.html) string (or an Excel formula).</span></span> <span data-ttu-id="647b0-155">A seguir, há um exemplo que define valores válidos como datas na primeira semana de abril de 2018.</span><span class="sxs-lookup"><span data-stu-id="647b0-155">The following is an example that defines valid values as dates in the first week of April, 2018.</span></span> 

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

#### <a name="list-validation-rule-type"></a><span data-ttu-id="647b0-156">Tipo de regra de validação de lista</span><span class="sxs-lookup"><span data-stu-id="647b0-156">List validation rule type</span></span>

<span data-ttu-id="647b0-157">Use a propriedade `list` no objeto `DataValidationRule` para especificar que os únicos valores válidos sejam aqueles de uma lista finita.</span><span class="sxs-lookup"><span data-stu-id="647b0-157">Use the `list` property in the `DataValidationRule` object to specify that the only valid values are those from a finite list.</span></span> <span data-ttu-id="647b0-158">Apresentamos um exemplo a seguir.</span><span class="sxs-lookup"><span data-stu-id="647b0-158">The following is an example.</span></span> <span data-ttu-id="647b0-159">Observe o seguinte sobre este código:</span><span class="sxs-lookup"><span data-stu-id="647b0-159">Note the following about this code:</span></span>

- <span data-ttu-id="647b0-160">Ele pressupõe que há uma planilha chamada “Nomes” e que os valores no intervalo “A1: A3” são nomes.</span><span class="sxs-lookup"><span data-stu-id="647b0-160">It assumes that there is a worksheet named "Names" and that the values in the range "A1:A3" are names.</span></span>
- <span data-ttu-id="647b0-161">A propriedade `source` especifica a lista de valores válidos.</span><span class="sxs-lookup"><span data-stu-id="647b0-161">The `source` property specifies the list of valid values.</span></span> <span data-ttu-id="647b0-162">O intervalo com os nomes foi atribuído a ela.</span><span class="sxs-lookup"><span data-stu-id="647b0-162">The range with the names has been assigned to it.</span></span> <span data-ttu-id="647b0-163">Também é possível atribuir uma lista delimitada por vírgula, por exemplo: “Sue, Ricky, Liz”.</span><span class="sxs-lookup"><span data-stu-id="647b0-163">You can also assign a comma-delimited list; for example: "Sue, Ricky, Liz".</span></span> 
- <span data-ttu-id="647b0-164">A propriedade `inCellDropDown` especifica se um controle suspenso aparecerá na célula quando o usuário selecioná-lo.</span><span class="sxs-lookup"><span data-stu-id="647b0-164">The `inCellDropDown` property specifies whether a drop-down control will appear in the cell when the user selects it.</span></span> <span data-ttu-id="647b0-165">Se definida como `true`, a lista suspensa aparece com a lista de valores de `source`.</span><span class="sxs-lookup"><span data-stu-id="647b0-165">If set to `true`, then the drop-down appears with the list of values from the `source`.</span></span>

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

#### <a name="custom-validation-rule-type"></a><span data-ttu-id="647b0-166">Tipo de regra de validação personalizada</span><span class="sxs-lookup"><span data-stu-id="647b0-166">Custom validation rule type</span></span>

<span data-ttu-id="647b0-167">Use a propriedade `custom` no objeto `DataValidationRule` para especificar uma fórmula de validação personalizada.</span><span class="sxs-lookup"><span data-stu-id="647b0-167">Use the `custom` property in the `DataValidationRule` object to specify a custom validation formula.</span></span> <span data-ttu-id="647b0-168">Apresentamos um exemplo a seguir.</span><span class="sxs-lookup"><span data-stu-id="647b0-168">The following is an example.</span></span> <span data-ttu-id="647b0-169">Observe o seguinte sobre este código:</span><span class="sxs-lookup"><span data-stu-id="647b0-169">Note the following about this code:</span></span>

- <span data-ttu-id="647b0-170">Ele pressupõe que há uma tabela de duas colunas com colunas **Nome do Atleta** e **Comentários** nas colunas A e B da planilha.</span><span class="sxs-lookup"><span data-stu-id="647b0-170">It assumes there is a two-column table with columns **Athlete Name** and **Comments** in the A and B columns of the worksheet.</span></span>
- <span data-ttu-id="647b0-171">Para reduzir a verbosidade na coluna **Comentários**, ele faz com que os dados que incluem o nome do atleta se tornem inválidos.</span><span class="sxs-lookup"><span data-stu-id="647b0-171">To reduce verbosity in the **Comments** column, it makes data that includes the athlete's name invalid.</span></span>
- <span data-ttu-id="647b0-172">`SEARCH(A2,B2)` retorna a posição inicial, na sequência de caracteres B2, da sequência de caracteres em A2.</span><span class="sxs-lookup"><span data-stu-id="647b0-172">`SEARCH(A2,B2)` returns the starting position, in string in B2, of the string in A2.</span></span> <span data-ttu-id="647b0-173">Se A2 não estiver contida em B2, ele não retornará um número.</span><span class="sxs-lookup"><span data-stu-id="647b0-173">If A2 is not contained in B2, it does not return a number.</span></span> <span data-ttu-id="647b0-174">`ISNUMBER()` retorna um booleano.</span><span class="sxs-lookup"><span data-stu-id="647b0-174">Returns a `ISNUMBER()`.</span></span> <span data-ttu-id="647b0-175">Então a propriedade `formula` diz que dados válidos da coluna **Comentários** são dados que não incluem a sequência de caracteres na coluna **Nome do Atleta**.</span><span class="sxs-lookup"><span data-stu-id="647b0-175">So the `formula` property says that valid data for the **Comment** column is data that does not include the string in the **Athlete Name** column.</span></span>

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

### <a name="create-validation-error-alerts"></a><span data-ttu-id="647b0-176">Criar alertas de erro de validação</span><span class="sxs-lookup"><span data-stu-id="647b0-176">Create validation error alerts</span></span>

<span data-ttu-id="647b0-177">É possível criar um alerta de erro personalizado que aparecerá quando um usuário tentar inserir dados inválidos em uma célula.</span><span class="sxs-lookup"><span data-stu-id="647b0-177">You can a create custom error alert that appears when a user tries to enter invalid data in a cell.</span></span> <span data-ttu-id="647b0-178">Há um exemplo simples a seguir.</span><span class="sxs-lookup"><span data-stu-id="647b0-178">The following is a simple example:</span></span> <span data-ttu-id="647b0-179">Observe o seguinte sobre este código:</span><span class="sxs-lookup"><span data-stu-id="647b0-179">Note the following about this code:</span></span>

- <span data-ttu-id="647b0-180">A propriedade `style` determina se o usuário recebe um alerta informativo, um aviso ou um alerta do tipo “pare”.</span><span class="sxs-lookup"><span data-stu-id="647b0-180">The `style` property determines whether the user gets an informational alert, a warning, or a "stop" alert.</span></span> <span data-ttu-id="647b0-181">Somente `Stop` impede de verdade que o usuário adicione dados inválidos.</span><span class="sxs-lookup"><span data-stu-id="647b0-181">Only `Stop` actually prevents the user from adding invalid data.</span></span> <span data-ttu-id="647b0-182">O pop-up para `Warning` e `Information` tem opções que permitem que o usuário insira os dados inválidos.</span><span class="sxs-lookup"><span data-stu-id="647b0-182">The pop-up for `Warning` and `Information` has options that allow the user enter the invalid data anyway.</span></span>
- <span data-ttu-id="647b0-183">A propriedade `showAlert` se torna padrão para `true`.</span><span class="sxs-lookup"><span data-stu-id="647b0-183">The `showAlert` property defaults to `true`.</span></span> <span data-ttu-id="647b0-184">Isso significa que o host do Excel exibirá um alerta pop-up genérico (do tipo `Stop`) a menos que seja criado um alerta personalizado que defina `showAlert` para `false` ou defina uma mensagem, um título e um estilo personalizados.</span><span class="sxs-lookup"><span data-stu-id="647b0-184">This means that the Excel host will pop-up a generic alert (of type `Stop`) unless you create a custom alert which either sets `showAlert` to `false` or sets a custom message, title, and style.</span></span> <span data-ttu-id="647b0-185">Esse código define uma mensagem personalizada e um título.</span><span class="sxs-lookup"><span data-stu-id="647b0-185">This code sets a custom message and title.</span></span>


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

<span data-ttu-id="647b0-186">Para obter mais informações, confira [DataValidationErrorAlert](https://docs.microsoft.com/javascript/api/excel/excel.datavalidationerroralert).</span><span class="sxs-lookup"><span data-stu-id="647b0-186">For more information, see [](https://docs.microsoft.com/javascript/api/excel/excel.datavalidationerroralert).</span></span>

### <a name="create-validation-prompts"></a><span data-ttu-id="647b0-187">Criar solicitações de validação</span><span class="sxs-lookup"><span data-stu-id="647b0-187">Create validation prompts</span></span>

<span data-ttu-id="647b0-188">É possível criar uma solicitação de instrução que aparece quando um usuário seleciona ou passa o mouse sobre uma célula na qual a validação de dados foi aplicada.</span><span class="sxs-lookup"><span data-stu-id="647b0-188">You can create an instructional prompt that appears when a user hovers over, or selects, a cell to which data validation has been applied.</span></span> <span data-ttu-id="647b0-189">Este é um exemplo:</span><span class="sxs-lookup"><span data-stu-id="647b0-189">The following is an example:</span></span>

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

<span data-ttu-id="647b0-190">Para obter mais informações, confira [DataValidationPrompt](https://docs.microsoft.com/javascript/api/excel/excel.datavalidationprompt).</span><span class="sxs-lookup"><span data-stu-id="647b0-190">For more information, see [](https://docs.microsoft.com/javascript/api/excel/excel.datavalidationprompt).</span></span>

### <a name="remove-data-validation-from-a-range"></a><span data-ttu-id="647b0-191">Remover a validação de dados de um intervalo</span><span class="sxs-lookup"><span data-stu-id="647b0-191">Remove data validation from a range</span></span>

<span data-ttu-id="647b0-192">Para remover a validação de dados de um intervalo, chame o método [Range.dataValidation.clear()](https://docs.microsoft.com/javascript/api/excel/excel.datavalidation#clear).</span><span class="sxs-lookup"><span data-stu-id="647b0-192">To remove data validation from a range, call the  [Range.dataValidation.clear()](https://docs.microsoft.com/javascript/api/excel/excel.datavalidation#clear) method.</span></span>

```js
myrange.dataValidation.clear()
```

<span data-ttu-id="647b0-193">Não é necessário que o intervalo limpo seja exatamente o mesmo de um intervalo no qual a validação de dados foi adicionada.</span><span class="sxs-lookup"><span data-stu-id="647b0-193">It isn't necessary that the range you clear is exactly the same range as a range on which you added data validation.</span></span> <span data-ttu-id="647b0-194">Se não for, apenas as células sobrepostas, se houver, dos dois intervalos serão limpas.</span><span class="sxs-lookup"><span data-stu-id="647b0-194">If it isn't, only the overlapping cells, if any, of the two ranges are cleared.</span></span> 

> [!NOTE]
> <span data-ttu-id="647b0-195">A limpeza da validação de dados de um intervalo também limpará qualquer validação de dados que um usuário tenha adicionado manualmente ao intervalo.</span><span class="sxs-lookup"><span data-stu-id="647b0-195">Clearing data validation from a range will also clear any data validation that a user has added manually to the range.</span></span>

## <a name="see-also"></a><span data-ttu-id="647b0-196">Veja também</span><span class="sxs-lookup"><span data-stu-id="647b0-196">See also</span></span>

- [<span data-ttu-id="647b0-197">Principais conceitos da API JavaScript do Excel</span><span class="sxs-lookup"><span data-stu-id="647b0-197">Excel JavaScript API core concepts</span></span>](excel-add-ins-core-concepts.md)
- [<span data-ttu-id="647b0-198">Objeto DataValidation (API JavaScript para Excel)</span><span class="sxs-lookup"><span data-stu-id="647b0-198">Chart Object (JavaScript API for Excel)</span></span>](https://docs.microsoft.com/javascript/api/excel/excel.datavalidation)
- [<span data-ttu-id="647b0-199">Objeto Range (API JavaScript para Excel)</span><span class="sxs-lookup"><span data-stu-id="647b0-199">Range Object (JavaScript API for Excel)</span></span>](https://docs.microsoft.com/javascript/api/excel/excel.range)



 
