---
title: Adicionar validação de dados para intervalos do Excel
description: Saiba como as EXCEL JavaScript permitem que seu complemento adicione validação automática de dados a tabelas, colunas, linhas e outros intervalos em uma workbook.
ms.date: 03/19/2019
localization_priority: Normal
ms.openlocfilehash: e1f5729e6e85ff8af92968c2ad65c19e655106e2
ms.sourcegitcommit: 883f71d395b19ccfc6874a0d5942a7016eb49e2c
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 07/09/2021
ms.locfileid: "53349522"
---
# <a name="add-data-validation-to-excel-ranges"></a><span data-ttu-id="6ee44-103">Adicionar validação de dados para intervalos do Excel</span><span class="sxs-lookup"><span data-stu-id="6ee44-103">Add data validation to Excel ranges</span></span>

<span data-ttu-id="6ee44-104">A biblioteca JavaScript do Excel fornece APIs para permitir que seu suplemento adicione a validação de dados automáticos a tabelas, colunas, linhas e outros intervalos em uma pasta de trabalho.</span><span class="sxs-lookup"><span data-stu-id="6ee44-104">The Excel JavaScript Library provides APIs to enable your add-in to add automatic data validation to tables, columns, rows, and other ranges in a workbook.</span></span> <span data-ttu-id="6ee44-105">Para entender os conceitos e a terminologia da validação de dados, consulte os artigos a seguir sobre como os usuários adicionam validação de dados por meio da interface do usuário Excel usuário.</span><span class="sxs-lookup"><span data-stu-id="6ee44-105">To understand the concepts and the terminology of data validation, please see the following articles about how users add data validation through the Excel UI.</span></span>

- [<span data-ttu-id="6ee44-106">Aplicar validação de dados às células</span><span class="sxs-lookup"><span data-stu-id="6ee44-106">Apply data validation to cells</span></span>](https://support.office.com/article/Apply-data-validation-to-cells-29FECBCC-D1B9-42C1-9D76-EFF3CE5F7249)
- [<span data-ttu-id="6ee44-107">Validação de dados</span><span class="sxs-lookup"><span data-stu-id="6ee44-107">More on data validation</span></span>](https://support.office.com/article/More-on-data-validation-f38dee73-9900-4ca6-9301-8a5f6e1f0c4c)
- [<span data-ttu-id="6ee44-108">Exemplos e descrição de validação de dados no Excel</span><span class="sxs-lookup"><span data-stu-id="6ee44-108">Description and examples of data validation in Excel</span></span>](https://support.microsoft.com/help/211485/description-and-examples-of-data-validation-in-excel)

## <a name="programmatic-control-of-data-validation"></a><span data-ttu-id="6ee44-109">Controle de programação de validação de dados</span><span class="sxs-lookup"><span data-stu-id="6ee44-109">Programmatic control of data validation</span></span>

<span data-ttu-id="6ee44-110">A `Range.dataValidation` propriedade, que usa um objeto [DataValidation](/javascript/api/excel/excel.datavalidation), é o ponto de entrada para o controle de programação de validação de dados no Excel.</span><span class="sxs-lookup"><span data-stu-id="6ee44-110">The `Range.dataValidation` property, which takes a [DataValidation](/javascript/api/excel/excel.datavalidation) object, is the entry point for programmatic control of data validation in Excel.</span></span> <span data-ttu-id="6ee44-111">Há cinco propriedades para o objeto `DataValidation`:</span><span class="sxs-lookup"><span data-stu-id="6ee44-111">There are five properties to the `DataValidation` object:</span></span>

- <span data-ttu-id="6ee44-112">`rule` &#8212;Define o que constitui dados válidos para o intervalo.</span><span class="sxs-lookup"><span data-stu-id="6ee44-112">`rule` &#8212; Defines what constitutes valid data for the range.</span></span> <span data-ttu-id="6ee44-113">Ver [DataValidationRule](/javascript/api/excel/excel.datavalidationrule).</span><span class="sxs-lookup"><span data-stu-id="6ee44-113">See [DataValidationRule](/javascript/api/excel/excel.datavalidationrule).</span></span>
- <span data-ttu-id="6ee44-114">`errorAlert` &#8212;Especifica se um erro é exibido se o usuário insere dados inválidos e define o texto, o título e o estilo de alerta; Por exemplo, **informativo**, **Aviso**, e **Parar**.</span><span class="sxs-lookup"><span data-stu-id="6ee44-114">`errorAlert` &#8212; Specifies whether an error pops up if the user enters invalid data, and defines the alert text, title, and style; for example, **Informational**, **Warning**, and **Stop**.</span></span> <span data-ttu-id="6ee44-115">Ver [DataValidationErrorAlert](/javascript/api/excel/excel.datavalidationerroralert).</span><span class="sxs-lookup"><span data-stu-id="6ee44-115">See [DataValidationErrorAlert](/javascript/api/excel/excel.datavalidationerroralert).</span></span>
- <span data-ttu-id="6ee44-116">`prompt` &#8212;Especifica se um aviso aparece quando o usuário passa o mouse sobre o intervalo e define a mensagem de aviso.</span><span class="sxs-lookup"><span data-stu-id="6ee44-116">`prompt` &#8212; Specifies whether a prompt appears when the user hovers over the range and defines the prompt message.</span></span> <span data-ttu-id="6ee44-117">Ver [DataValidationPrompt](/javascript/api/excel/excel.datavalidationprompt).</span><span class="sxs-lookup"><span data-stu-id="6ee44-117">See [DataValidationPrompt](/javascript/api/excel/excel.datavalidationprompt).</span></span>
- <span data-ttu-id="6ee44-118">`ignoreBlanks` &#8212;Especifica se aplica a regra de validação de dados a células em branco no intervalo.</span><span class="sxs-lookup"><span data-stu-id="6ee44-118">`ignoreBlanks` &#8212; Specifies whether the data validation rule applies to blank cells in the range.</span></span> <span data-ttu-id="6ee44-119">O padrão é `true`</span><span class="sxs-lookup"><span data-stu-id="6ee44-119">Defaults to `true`.</span></span>
- <span data-ttu-id="6ee44-120">`type` &#8212;Identificação somente leitura do tipo de validação, como WholeNumber, data, TextLength etc. Ela é definida indiretamente quando você define a propriedade `rule`.</span><span class="sxs-lookup"><span data-stu-id="6ee44-120">`type` &#8212; A read-only identification of the validation type, such as WholeNumber, Date, TextLength, etc. It is set indirectly when you set the `rule` property.</span></span>

> [!NOTE]
> <span data-ttu-id="6ee44-121">A validação de dados adicionada programaticamente funciona exatamente como a validação de dados adicionada manualmente.</span><span class="sxs-lookup"><span data-stu-id="6ee44-121">Data validation added programmatically behaves just like manually added data validation.</span></span> <span data-ttu-id="6ee44-122">Em particular, observe que a validação de dados é disparada somente se o usuário inserir diretamente um valor em uma célula ou copiar e colar uma célula de outro local da pasta de trabalho e escolher a opção de colagem **Valores**.</span><span class="sxs-lookup"><span data-stu-id="6ee44-122">In particular, note that data validation is triggered only if the user directly enters a value into a cell or copies and pastes a cell from elsewhere in the workbook and chooses the **Values** paste option.</span></span> <span data-ttu-id="6ee44-123">Se o usuário copiar uma célula e fizer uma colagem simples em um intervalo com a validação de dados, a validação não será disparada.</span><span class="sxs-lookup"><span data-stu-id="6ee44-123">If the user copies a cell and does a plain paste into a range with data validation, validation is not triggered.</span></span>

## <a name="creating-validation-rules"></a><span data-ttu-id="6ee44-124">Criar regras de validação</span><span class="sxs-lookup"><span data-stu-id="6ee44-124">Creating validation rules</span></span>

<span data-ttu-id="6ee44-125">Para adicionar a validação de dados a um intervalo, o código deve configurar a propriedade `rule` do objeto `DataValidation` em `Range.dataValidation`.</span><span class="sxs-lookup"><span data-stu-id="6ee44-125">To add data validation to a range, your code must set the `rule` property of the `DataValidation` object in `Range.dataValidation`.</span></span> <span data-ttu-id="6ee44-126">Isso leva ao objeto [DataValidationRule](/javascript/api/excel/excel.datavalidationrule) que tem sete propriedades opcionais.</span><span class="sxs-lookup"><span data-stu-id="6ee44-126">This takes a [DataValidationRule](/javascript/api/excel/excel.datavalidationrule) object which has seven optional properties.</span></span> <span data-ttu-id="6ee44-127">*Não mais de uma dessas propriedades pode estar presente em qualquer objeto `DataValidationRule`.*</span><span class="sxs-lookup"><span data-stu-id="6ee44-127">*No more than one of these properties may be present in any `DataValidationRule` object.*</span></span> <span data-ttu-id="6ee44-128">A propriedade que você incluir determina o tipo de validação.</span><span class="sxs-lookup"><span data-stu-id="6ee44-128">The property that you include determines the type of validation.</span></span>

### <a name="basic-and-datetime-validation-rule-types"></a><span data-ttu-id="6ee44-129">Tipos de regras de validação Basic e DateTime</span><span class="sxs-lookup"><span data-stu-id="6ee44-129">Basic and DateTime validation rule types</span></span>

<span data-ttu-id="6ee44-130">As três primeiras propriedades `DataValidationRule` (ou seja, tipos de regra de validação) consideram o objeto [BasicDataValidation](/javascript/api/excel/excel.basicdatavalidation) como o seu valor.</span><span class="sxs-lookup"><span data-stu-id="6ee44-130">The first three `DataValidationRule` properties (i.e., validation rule types) take a [BasicDataValidation](/javascript/api/excel/excel.basicdatavalidation) object as their value.</span></span>

- <span data-ttu-id="6ee44-131">`wholeNumber` &#8212;Requer um número inteiro, além de outra validação especificada pelo objeto `BasicDataValidation`.</span><span class="sxs-lookup"><span data-stu-id="6ee44-131">`wholeNumber` &#8212; Requires a whole number in addition to any other validation specified by the `BasicDataValidation` object.</span></span>
- <span data-ttu-id="6ee44-132">`decimal` &#8212;Requer um número decimal, além de outra validação especificada pelo objeto `BasicDataValidation`.</span><span class="sxs-lookup"><span data-stu-id="6ee44-132">`decimal` &#8212; Requires a decimal number in addition to any other validation specified by the `BasicDataValidation` object.</span></span>
- <span data-ttu-id="6ee44-133">`textLength` &#8212;Aplicam-se os detalhes de validação do objeto `BasicDataValidation` ao *comprimento* de valor da célula.</span><span class="sxs-lookup"><span data-stu-id="6ee44-133">`textLength` &#8212; Applies the validation details in the `BasicDataValidation` object to the *length* of the cell's value.</span></span>

<span data-ttu-id="6ee44-134">Aqui está um exemplo de como criar uma regra de validação.</span><span class="sxs-lookup"><span data-stu-id="6ee44-134">Here is an example of creating a validation rule.</span></span> <span data-ttu-id="6ee44-135">Observe o seguinte sobre este código.</span><span class="sxs-lookup"><span data-stu-id="6ee44-135">Note the following about this code.</span></span>

- <span data-ttu-id="6ee44-136">O `operator` é o operador binário "GreaterThan".</span><span class="sxs-lookup"><span data-stu-id="6ee44-136">The `operator` is the binary operator "GreaterThan".</span></span> <span data-ttu-id="6ee44-137">Sempre que você usa um operador binário, o valor que o usuário tenta inserir na célula é o operando à esquerda e o valor especificado em `formula1` é o operando à direita.</span><span class="sxs-lookup"><span data-stu-id="6ee44-137">Whenever you use a binary operator, the value that the user tries to enter in the cell is the left-hand operand and the value specified in `formula1` is the right-hand operand.</span></span> <span data-ttu-id="6ee44-138">Então esta regra diz que apenas números inteiros que são maiores do que 0 são válidos.</span><span class="sxs-lookup"><span data-stu-id="6ee44-138">So this rule says that only whole numbers that are greater than 0 are valid.</span></span> 
- <span data-ttu-id="6ee44-139">O `formula1` é um número embutido.</span><span class="sxs-lookup"><span data-stu-id="6ee44-139">The `formula1` is a hard-coded number.</span></span> <span data-ttu-id="6ee44-140">Se não souber no momento da codificação qual é o valor, você também poderá usar uma fórmula do Excel (como uma cadeia de caracteres) para o valor.</span><span class="sxs-lookup"><span data-stu-id="6ee44-140">If you don't know at coding time what the value should be, you can also use an Excel formula (as a string) for the value.</span></span> <span data-ttu-id="6ee44-141">Por exemplo, "= A3" e "SOMA(A4,B5) =" também seriam valores `formula1`.</span><span class="sxs-lookup"><span data-stu-id="6ee44-141">For example, "=A3" and "=SUM(A4,B5)" could also be values of `formula1`.</span></span>

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

<span data-ttu-id="6ee44-142">Confira [BasicDataValidation](/javascript/api/excel/excel.basicdatavalidation) para uma lista de outros operadores binários.</span><span class="sxs-lookup"><span data-stu-id="6ee44-142">See [BasicDataValidation](/javascript/api/excel/excel.basicdatavalidation) for a list of the other binary operators.</span></span> 

<span data-ttu-id="6ee44-143">Também há dois operadores ternários: "Between" e "NotBetween".</span><span class="sxs-lookup"><span data-stu-id="6ee44-143">There are also two ternary operators: "Between" and "NotBetween".</span></span> <span data-ttu-id="6ee44-144">Para usá-los, você deve especificar a propriedade `formula2` opcional.</span><span class="sxs-lookup"><span data-stu-id="6ee44-144">To use these, you must specify the optional `formula2` property.</span></span> <span data-ttu-id="6ee44-145">Os valores`formula1` e `formula2` são os operandos delimitadores.</span><span class="sxs-lookup"><span data-stu-id="6ee44-145">The `formula1` and `formula2` values are the bounding operands.</span></span> <span data-ttu-id="6ee44-146">O valor que o usuário tenta inserir na célula é o terceiro operando (calculado).</span><span class="sxs-lookup"><span data-stu-id="6ee44-146">The value that the user tries to enter in the cell is the third (evaluated) operand.</span></span> <span data-ttu-id="6ee44-147">A seguir, um exemplo de uso do operador "Between".</span><span class="sxs-lookup"><span data-stu-id="6ee44-147">The following is an example of using the "Between" operator.</span></span>

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

<span data-ttu-id="6ee44-148">As próximas duas regras de propriedades usam o objeto [DateTimeDataValidation](/javascript/api/excel/excel.datetimedatavalidation) como seu valor.</span><span class="sxs-lookup"><span data-stu-id="6ee44-148">The next two rule properties take a [DateTimeDataValidation](/javascript/api/excel/excel.datetimedatavalidation) object as their value.</span></span>

- `date`
- `time`

<span data-ttu-id="6ee44-149">O objeto `DateTimeDataValidation` é estruturado da mesma forma que o `BasicDataValidation`: com as propriedades `formula1`, `formula2` e `operator`, e é usado da mesma maneira.</span><span class="sxs-lookup"><span data-stu-id="6ee44-149">The `DateTimeDataValidation` object is structured similarly to the `BasicDataValidation`: it has the properties `formula1`, `formula2`, and `operator`, and is used in the same way.</span></span> <span data-ttu-id="6ee44-150">A diferença é que você não pode usar um número nas propriedades de fórmula, mas você pode inserir uma cadeia [ISO 8606 datetime](https://www.iso.org/iso-8601-date-and-time-format.html) (ou uma fórmula do Excel).</span><span class="sxs-lookup"><span data-stu-id="6ee44-150">The difference is that you cannot use a number in the formula properties, but you can enter a [ISO 8606 datetime](https://www.iso.org/iso-8601-date-and-time-format.html) string (or an Excel formula).</span></span> <span data-ttu-id="6ee44-151">A seguir está um exemplo que define os valores válidos como datas na primeira semana de abril de 2018.</span><span class="sxs-lookup"><span data-stu-id="6ee44-151">The following is an example that defines valid values as dates in the first week of April, 2018.</span></span> 

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

### <a name="list-validation-rule-type"></a><span data-ttu-id="6ee44-152">Tipos de regra de validação de lista</span><span class="sxs-lookup"><span data-stu-id="6ee44-152">List validation rule type</span></span>

<span data-ttu-id="6ee44-153">Use a propriedade `list` do objeto `DataValidationRule` para especificar valores que são válidos apenas em uma lista finita.</span><span class="sxs-lookup"><span data-stu-id="6ee44-153">Use the `list` property in the `DataValidationRule` object to specify that the only valid values are those from a finite list.</span></span> <span data-ttu-id="6ee44-154">Apresentamos um exemplo a seguir.</span><span class="sxs-lookup"><span data-stu-id="6ee44-154">The following is an example.</span></span> <span data-ttu-id="6ee44-155">Observe o seguinte sobre este código.</span><span class="sxs-lookup"><span data-stu-id="6ee44-155">Note the following about this code.</span></span>

- <span data-ttu-id="6ee44-156">Ele pressupõe que se trata de uma planilha chamada "Nomes" e que os valores no intervalo "A1: A3" são nomes.</span><span class="sxs-lookup"><span data-stu-id="6ee44-156">It assumes that there is a worksheet named "Names" and that the values in the range "A1:A3" are names.</span></span>
- <span data-ttu-id="6ee44-157">A propriedade `source` especifica a lista de valores válidos.</span><span class="sxs-lookup"><span data-stu-id="6ee44-157">The `source` property specifies the list of valid values.</span></span> <span data-ttu-id="6ee44-158">O argumento de cadeia de caracteres se refere a um intervalo que contém os nomes.</span><span class="sxs-lookup"><span data-stu-id="6ee44-158">The string argument refers to a range containing the names.</span></span> <span data-ttu-id="6ee44-159">Você também pode atribuir uma lista delimitada por vírgula; por exemplo: "Lara, Pedro, Marina".</span><span class="sxs-lookup"><span data-stu-id="6ee44-159">You can also assign a comma-delimited list; for example: "Sue, Ricky, Liz".</span></span> 
- <span data-ttu-id="6ee44-160">A propriedade `inCellDropDown` especifica se um controle de lista suspensa será exibido na célula quando o usuário a selecionar.</span><span class="sxs-lookup"><span data-stu-id="6ee44-160">The `inCellDropDown` property specifies whether a drop-down control will appear in the cell when the user selects it.</span></span> <span data-ttu-id="6ee44-161">Se definido como `true`, em seguida, a lista suspensa é exibida com a lista de valores do `source`.</span><span class="sxs-lookup"><span data-stu-id="6ee44-161">If set to `true`, then the drop-down appears with the list of values from the `source`.</span></span>

```js
Excel.run(function (context) {
    var sheet = context.workbook.worksheets.getActiveWorksheet();
    var range = sheet.getRange("B2:C5");   
    var nameSourceRange = context.workbook.worksheets.getItem("Names").getRange("A1:A3");

    range.dataValidation.rule = {
        list: {
            inCellDropDown: true,
            source: "=Names!$A$1:$A$3"
        }
    };

    return context.sync();
})
```

### <a name="custom-validation-rule-type"></a><span data-ttu-id="6ee44-162">Tipo de regra de validação personalizada</span><span class="sxs-lookup"><span data-stu-id="6ee44-162">Custom validation rule type</span></span>

<span data-ttu-id="6ee44-163">Use a propriedade `custom` no objeto `DataValidationRule` para especificar uma fórmula de validação personalizada.</span><span class="sxs-lookup"><span data-stu-id="6ee44-163">Use the `custom` property in the `DataValidationRule` object to specify a custom validation formula.</span></span> <span data-ttu-id="6ee44-164">Apresentamos um exemplo a seguir.</span><span class="sxs-lookup"><span data-stu-id="6ee44-164">The following is an example.</span></span> <span data-ttu-id="6ee44-165">Observe o seguinte sobre este código.</span><span class="sxs-lookup"><span data-stu-id="6ee44-165">Note the following about this code.</span></span>

- <span data-ttu-id="6ee44-166">Ele pressupõe que há uma tabela de duas colunas com as colunas **Nome do Atleta** e **Comentários** nas colunas A e B da planilha.</span><span class="sxs-lookup"><span data-stu-id="6ee44-166">It assumes there is a two-column table with columns **Athlete Name** and **Comments** in the A and B columns of the worksheet.</span></span>
- <span data-ttu-id="6ee44-167">Para reduzir o nível de detalhamento na coluna **Comentários**, ela torna os dados que incluem o nome do atleta inválidos.</span><span class="sxs-lookup"><span data-stu-id="6ee44-167">To reduce verbosity in the **Comments** column, it makes data that includes the athlete's name invalid.</span></span>
- <span data-ttu-id="6ee44-168">`SEARCH(A2,B2)` Retorna a posição inicial, na cadeia de caracteres em B2, da cadeia de caracteres em A2.</span><span class="sxs-lookup"><span data-stu-id="6ee44-168">`SEARCH(A2,B2)` returns the starting position, in string in B2, of the string in A2.</span></span> <span data-ttu-id="6ee44-169">Se A2 não estiver contida em B2, ela não retornará um número.</span><span class="sxs-lookup"><span data-stu-id="6ee44-169">If A2 is not contained in B2, it does not return a number.</span></span> <span data-ttu-id="6ee44-170">`ISNUMBER()` retorna um booliano.</span><span class="sxs-lookup"><span data-stu-id="6ee44-170">`ISNUMBER()` returns a boolean.</span></span> <span data-ttu-id="6ee44-171">Portanto, a propriedade `formula` diz que os dados válidos para a coluna **Comentário** são os dados que não incluem a cadeia de caracteres da coluna **Nome do Atleta**.</span><span class="sxs-lookup"><span data-stu-id="6ee44-171">So the `formula` property says that valid data for the **Comment** column is data that does not include the string in the **Athlete Name** column.</span></span>

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

## <a name="create-validation-error-alerts"></a><span data-ttu-id="6ee44-172">Criar alertas de erro de validação</span><span class="sxs-lookup"><span data-stu-id="6ee44-172">Create validation error alerts</span></span>

<span data-ttu-id="6ee44-173">Você pode criar um alerta de erro personalizado que aparece quando um usuário tenta inserir dados inválidos em uma célula.</span><span class="sxs-lookup"><span data-stu-id="6ee44-173">You can a create custom error alert that appears when a user tries to enter invalid data in a cell.</span></span> <span data-ttu-id="6ee44-174">Apresentamos um exemplo simples a seguir.</span><span class="sxs-lookup"><span data-stu-id="6ee44-174">The following is a simple example.</span></span> <span data-ttu-id="6ee44-175">Observe o seguinte sobre este código.</span><span class="sxs-lookup"><span data-stu-id="6ee44-175">Note the following about this code.</span></span>

- <span data-ttu-id="6ee44-176">A propriedade `style` determina se o usuário recebe um alerta informativo, um aviso e um alerta "parar".</span><span class="sxs-lookup"><span data-stu-id="6ee44-176">The `style` property determines whether the user gets an informational alert, a warning, or a "stop" alert.</span></span> <span data-ttu-id="6ee44-177">Apenas `Stop` realmente impede que o usuário adicione dados inválidos.</span><span class="sxs-lookup"><span data-stu-id="6ee44-177">Only `Stop` actually prevents the user from adding invalid data.</span></span> <span data-ttu-id="6ee44-178">O pop-up para `Warning` e `Information` tem opções para permitir que o usuário insira dados inválidos assim mesmo.</span><span class="sxs-lookup"><span data-stu-id="6ee44-178">The pop-up for `Warning` and `Information` has options that allow the user enter the invalid data anyway.</span></span>
- <span data-ttu-id="6ee44-179">As propriedades `showAlert` padrão para `true`.</span><span class="sxs-lookup"><span data-stu-id="6ee44-179">The `showAlert` property defaults to `true`.</span></span> <span data-ttu-id="6ee44-180">Isso significa Excel um alerta genérico (de tipo), a menos que você crie um alerta personalizado que define ou define uma mensagem, título e `Stop` `showAlert` estilo `false` personalizados.</span><span class="sxs-lookup"><span data-stu-id="6ee44-180">This means that Excel will pop-up a generic alert (of type `Stop`) unless you create a custom alert which either sets `showAlert` to `false` or sets a custom message, title, and style.</span></span> <span data-ttu-id="6ee44-181">O código define uma mensagem personalizada e o título.</span><span class="sxs-lookup"><span data-stu-id="6ee44-181">This code sets a custom message and title.</span></span>

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

<span data-ttu-id="6ee44-182">Para saber mais, confira [DataValidationErrorAlert](/javascript/api/excel/excel.datavalidationerroralert).</span><span class="sxs-lookup"><span data-stu-id="6ee44-182">For more information, see [DataValidationErrorAlert](/javascript/api/excel/excel.datavalidationerroralert).</span></span>

## <a name="create-validation-prompts"></a><span data-ttu-id="6ee44-183">Criar solicitações de validação</span><span class="sxs-lookup"><span data-stu-id="6ee44-183">Create validation prompts</span></span>

<span data-ttu-id="6ee44-184">Você pode criar um prompt instrutivo que é exibido quando um usuário passa o mouse sobre ele ou seleciona uma célula à qual os dados de validação foram aplicados.</span><span class="sxs-lookup"><span data-stu-id="6ee44-184">You can create an instructional prompt that appears when a user hovers over, or selects, a cell to which data validation has been applied.</span></span> <span data-ttu-id="6ee44-185">Apresentamos um exemplo a seguir.</span><span class="sxs-lookup"><span data-stu-id="6ee44-185">The following is an example.</span></span>

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

<span data-ttu-id="6ee44-186">Para saber mais, confira [DataValidationPrompt](/javascript/api/excel/excel.datavalidationprompt).</span><span class="sxs-lookup"><span data-stu-id="6ee44-186">For more information, see [DataValidationPrompt](/javascript/api/excel/excel.datavalidationprompt).</span></span>

## <a name="remove-data-validation-from-a-range"></a><span data-ttu-id="6ee44-187">Remover validação de dados de um intervalo</span><span class="sxs-lookup"><span data-stu-id="6ee44-187">Remove data validation from a range</span></span>

<span data-ttu-id="6ee44-188">Para remover a validação de dados de um intervalo, acione o método [Range.dataValidation.clear()](/javascript/api/excel/excel.datavalidation#clear--).</span><span class="sxs-lookup"><span data-stu-id="6ee44-188">To remove data validation from a range, call the  [Range.dataValidation.clear()](/javascript/api/excel/excel.datavalidation#clear--) method.</span></span>

```js
myrange.dataValidation.clear()
```

<span data-ttu-id="6ee44-189">Não é necessário que o intervalo que você desmarcar seja o mesmo intervalo de um intervalo no qual você adicionou a validação de dados.</span><span class="sxs-lookup"><span data-stu-id="6ee44-189">It isn't necessary that the range you clear is exactly the same range as a range on which you added data validation.</span></span> <span data-ttu-id="6ee44-190">Caso contrário, apenas as células sobrepostas, se houver, dos dois intervalos são desmarcadas.</span><span class="sxs-lookup"><span data-stu-id="6ee44-190">If it isn't, only the overlapping cells, if any, of the two ranges are cleared.</span></span> 

> [!NOTE]
> <span data-ttu-id="6ee44-191">Limpar a validação de dados de um intervalo também limpará qualquer validação de dados que o usuário tenha adicionado manualmente ao intervalo.</span><span class="sxs-lookup"><span data-stu-id="6ee44-191">Clearing data validation from a range will also clear any data validation that a user has added manually to the range.</span></span>

## <a name="see-also"></a><span data-ttu-id="6ee44-192">Confira também</span><span class="sxs-lookup"><span data-stu-id="6ee44-192">See also</span></span>

- [<span data-ttu-id="6ee44-193">Modelo de objeto JavaScript do Excel em Suplementos do Office</span><span class="sxs-lookup"><span data-stu-id="6ee44-193">Excel JavaScript object model in Office Add-ins</span></span>](excel-add-ins-core-concepts.md)
- [<span data-ttu-id="6ee44-194">Objeto Application (JavaScript API para Excel)</span><span class="sxs-lookup"><span data-stu-id="6ee44-194">DataValidation Object (JavaScript API for Excel)</span></span>](/javascript/api/excel/excel.datavalidation)
- [<span data-ttu-id="6ee44-195">Objeto Range (API JavaScript para Excel)</span><span class="sxs-lookup"><span data-stu-id="6ee44-195">Range Object (JavaScript API for Excel)</span></span>](/javascript/api/excel/excel.range)
