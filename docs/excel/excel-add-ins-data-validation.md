---
title: Adicionar validação de dados para intervalos do Excel
description: ''
ms.date: 10/03/2018
localization_priority: Priority
ms.openlocfilehash: dfe29bce5e23f7f251f6b52bf3eb359101f274ca
ms.sourcegitcommit: d1aa7201820176ed986b9f00bb9c88e055906c77
ms.translationtype: HT
ms.contentlocale: pt-BR
ms.lasthandoff: 01/23/2019
ms.locfileid: "29386985"
---
# <a name="add-data-validation-to-excel-ranges"></a><span data-ttu-id="7bb22-102">Adicionar validação de dados para intervalos do Excel</span><span class="sxs-lookup"><span data-stu-id="7bb22-102">Add data validation to Excel ranges</span></span>

<span data-ttu-id="7bb22-103">A biblioteca JavaScript do Excel fornece APIs para permitir que seu suplemento adicione a validação de dados automáticos a tabelas, colunas, linhas e outros intervalos em uma pasta de trabalho.</span><span class="sxs-lookup"><span data-stu-id="7bb22-103">The Excel JavaScript Library provides APIs to enable your add-in to add automatic data validation to tables, columns, rows, and other ranges in a workbook.</span></span> <span data-ttu-id="7bb22-104">Para entender os conceitos e a terminologia de validação de dados, confira os seguintes artigos sobre como os usuários adicionam a validação de dados na interface do usuário do Excel:</span><span class="sxs-lookup"><span data-stu-id="7bb22-104">To understand the concepts and the terminology of data validation, please see the following articles about how users add data validation through the Excel UI:</span></span>

- [<span data-ttu-id="7bb22-105">Apply data validation to cells</span><span class="sxs-lookup"><span data-stu-id="7bb22-105">Apply data validation to cells</span></span>](https://support.office.com/article/Apply-data-validation-to-cells-29FECBCC-D1B9-42C1-9D76-EFF3CE5F7249)
- [<span data-ttu-id="7bb22-106">Validação de dados</span><span class="sxs-lookup"><span data-stu-id="7bb22-106">More on data validation</span></span>](https://support.office.com/article/More-on-data-validation-f38dee73-9900-4ca6-9301-8a5f6e1f0c4c)
- [<span data-ttu-id="7bb22-107">Exemplos e descrição de validação de dados no Excel</span><span class="sxs-lookup"><span data-stu-id="7bb22-107">Description and examples of data validation in Excel</span></span>](https://support.microsoft.com/help/211485/description-and-examples-of-data-validation-in-excel)

## <a name="programmatic-control-of-data-validation"></a><span data-ttu-id="7bb22-108">Controle de programação de validação de dados</span><span class="sxs-lookup"><span data-stu-id="7bb22-108">Programmatic control of data validation</span></span>

<span data-ttu-id="7bb22-109">A `Range.dataValidation` propriedade, que usa um objeto [DataValidation](https://docs.microsoft.com/javascript/api/excel/excel.datavalidation), é o ponto de entrada para o controle de programação de validação de dados no Excel.</span><span class="sxs-lookup"><span data-stu-id="7bb22-109">The `Range.dataValidation` property, which takes a [DataValidation](https://docs.microsoft.com/javascript/api/excel/excel.datavalidation) object, is the entry point for programmatic control of data validation in Excel.</span></span> <span data-ttu-id="7bb22-110">Há cinco propriedades a `DataValidation` objeto:</span><span class="sxs-lookup"><span data-stu-id="7bb22-110">There are five properties to the `DataValidation` object:</span></span>

- <span data-ttu-id="7bb22-111">`rule` &#8212;Define o que constitui dados válidos para o intervalo.</span><span class="sxs-lookup"><span data-stu-id="7bb22-111">`rule` &#8212; Defines what constitutes valid data for the range.</span></span> <span data-ttu-id="7bb22-112">Ver [DataValidationRule](https://docs.microsoft.com/javascript/api/excel/excel.datavalidationrule).</span><span class="sxs-lookup"><span data-stu-id="7bb22-112">See [DataValidationRule](https://docs.microsoft.com/javascript/api/excel/excel.datavalidationrule).</span></span>
- <span data-ttu-id="7bb22-113">`errorAlert` &#8212;Especifica se um erro é exibido se o usuário insere dados inválidos e define o texto, o título e o estilo de alerta; Por exemplo, **informativo**, **Aviso**, e **Parar**.</span><span class="sxs-lookup"><span data-stu-id="7bb22-113">`errorAlert` &#8212; Specifies whether an error pops up if the user enters invalid data, and defines the alert text, title, and style; for example, **Informational**, **Warning**, and **Stop**.</span></span> <span data-ttu-id="7bb22-114">Ver [DataValidationErrorAlert](https://docs.microsoft.com/javascript/api/excel/excel.datavalidationerroralert).</span><span class="sxs-lookup"><span data-stu-id="7bb22-114">See [DataValidationErrorAlert](https://docs.microsoft.com/javascript/api/excel/excel.datavalidationerroralert).</span></span>
- <span data-ttu-id="7bb22-115">`prompt` &#8212;Especifica se um aviso aparece quando o usuário passa o mouse sobre o intervalo e define a mensagem de aviso.</span><span class="sxs-lookup"><span data-stu-id="7bb22-115">`prompt` &#8212; Specifies whether a prompt appears when the user hovers over the range and defines the prompt message.</span></span> <span data-ttu-id="7bb22-116">Ver [DataValidationPrompt](https://docs.microsoft.com/javascript/api/excel/excel.datavalidationprompt).</span><span class="sxs-lookup"><span data-stu-id="7bb22-116">See [DataValidationPrompt](https://docs.microsoft.com/javascript/api/excel/excel.datavalidationprompt).</span></span>
- <span data-ttu-id="7bb22-117">`ignoreBlanks` &#8212;Especifica se aplica a regra de validação de dados a células em branco no intervalo.</span><span class="sxs-lookup"><span data-stu-id="7bb22-117">`ignoreBlanks` &#8212; Specifies whether the data validation rule applies to blank cells in the range.</span></span> <span data-ttu-id="7bb22-118">O padrão é `true`</span><span class="sxs-lookup"><span data-stu-id="7bb22-118">Defaults to `true`.</span></span>
- <span data-ttu-id="7bb22-119">`type` &#8212;Identificação de somente leitura do tipo de validação, como WholeNumber, data, TextLength, etc. Ela é definida indiretamente quando você define a `rule` propriedade.</span><span class="sxs-lookup"><span data-stu-id="7bb22-119">`type` &#8212; A read-only identification of the validation type, such as WholeNumber, Date, TextLength, etc. It is set indirectly when you set the `rule` property.</span></span>

> [!NOTE]
> <span data-ttu-id="7bb22-120">A validação de dados adicionada programaticamente funciona exatamente como a validação de dados adicionada manualmente.</span><span class="sxs-lookup"><span data-stu-id="7bb22-120">Data validation added programmatically behaves just like manually added data validation.</span></span> <span data-ttu-id="7bb22-121">Em particular, observe que a validação de dados é disparada somente se o usuário inserir diretamente um valor em uma célula ou copiar e colar uma célula de outro local da pasta de trabalho e escolher os **valores** opção de colagem.</span><span class="sxs-lookup"><span data-stu-id="7bb22-121">In particular, note that data validation is triggered only if the user directly enters a value into a cell or copies and pastes a cell from elsewhere in the workbook and chooses the **Values** paste option.</span></span> <span data-ttu-id="7bb22-122">Se o usuário copiar uma célula e fazer uma colagem simples em um intervalo com a validação de dados, a validação não é disparada.</span><span class="sxs-lookup"><span data-stu-id="7bb22-122">If the user copies a cell and does a plain paste into a range with data validation, validation is not triggered.</span></span>

## <a name="creating-validation-rules"></a><span data-ttu-id="7bb22-123">Criar regras de validação</span><span class="sxs-lookup"><span data-stu-id="7bb22-123">Creating validation rules</span></span>

<span data-ttu-id="7bb22-124">Para adicionar a validação de dados em um intervalo, o código deve configurar a`rule` propriedade do `DataValidation` objeto em `Range.dataValidation`.</span><span class="sxs-lookup"><span data-stu-id="7bb22-124">To add data validation to a range, your code must set the `rule` property of the `DataValidation` object in `Range.dataValidation`.</span></span> <span data-ttu-id="7bb22-125">Isso leva ao objeto [DataValidationRule](https://docs.microsoft.com/javascript/api/excel/excel.datavalidationrule) que tem sete propriedades opcionais.</span><span class="sxs-lookup"><span data-stu-id="7bb22-125">This takes a [DataValidationRule](https://docs.microsoft.com/javascript/api/excel/excel.datavalidationrule) object which has seven optional properties.</span></span> <span data-ttu-id="7bb22-126">*Não mais de uma dessas propriedades pode estar presente em qualquer `DataValidationRule` objeto.*</span><span class="sxs-lookup"><span data-stu-id="7bb22-126">*No more than one of these properties may be present in any `DataValidationRule` object.*</span></span> <span data-ttu-id="7bb22-127">A propriedade que você incluir determina o tipo de validação.</span><span class="sxs-lookup"><span data-stu-id="7bb22-127">The property that you include determines the type of validation.</span></span>

### <a name="basic-and-datetime-validation-rule-types"></a><span data-ttu-id="7bb22-128">Tipos de regras de validação do Basic and DateTime</span><span class="sxs-lookup"><span data-stu-id="7bb22-128">Basic and DateTime validation rule types</span></span>

<span data-ttu-id="7bb22-129">As três primeiras `DataValidationRule` propriedades (ou seja, tipos de regra de validação) considere o objeto [BasicDataValidation](https://docs.microsoft.com/javascript/api/excel/excel.basicdatavalidation), como o valor.</span><span class="sxs-lookup"><span data-stu-id="7bb22-129">The first three `DataValidationRule` properties (i.e., validation rule types) take a [BasicDataValidation](https://docs.microsoft.com/javascript/api/excel/excel.basicdatavalidation) object as their value.</span></span>

- <span data-ttu-id="7bb22-130">`wholeNumber` &#8212;Requer um número inteiro, além de outra validação especificado pelo `BasicDataValidation` objeto.</span><span class="sxs-lookup"><span data-stu-id="7bb22-130">`wholeNumber` &#8212; Requires a whole number in addition to any other validation specified by the `BasicDataValidation` object.</span></span>
- <span data-ttu-id="7bb22-131">`decimal` &#8212;Requer um número decimal, além de outra validação especificada pelo `BasicDataValidation` objeto.</span><span class="sxs-lookup"><span data-stu-id="7bb22-131">`decimal` &#8212; Requires a decimal number in addition to any other validation specified by the `BasicDataValidation` object.</span></span>
- <span data-ttu-id="7bb22-132">`textLength` &#8212;Aplica-se os detalhes de validação no `BasicDataValidation` objeto para o *comprimento* de valor da célula.</span><span class="sxs-lookup"><span data-stu-id="7bb22-132">`textLength` &#8212; Applies the validation details in the `BasicDataValidation` object to the *length* of the cell's value.</span></span>

<span data-ttu-id="7bb22-133">Aqui está um exemplo de como criar uma regra de validação.</span><span class="sxs-lookup"><span data-stu-id="7bb22-133">Here is an example of creating a validation rule.</span></span> <span data-ttu-id="7bb22-134">Observe o seguinte sobre este código:</span><span class="sxs-lookup"><span data-stu-id="7bb22-134">Note the following about this code:</span></span>

- <span data-ttu-id="7bb22-135">O `operator` é o operador binário "GreaterThan".</span><span class="sxs-lookup"><span data-stu-id="7bb22-135">The `operator` is the binary operator "GreaterThan".</span></span> <span data-ttu-id="7bb22-136">Sempre que você usa um operador binário, o valor que o usuário tenta inserir na célula é operado à esquerda e o valor especificado no `formula1` é operado à direita.</span><span class="sxs-lookup"><span data-stu-id="7bb22-136">Whenever you use a binary operator, the value that the user tries to enter in the cell is the left-hand operand and the value specified in `formula1` is the right-hand operand.</span></span> <span data-ttu-id="7bb22-137">Então esta regra diz que apenas números inteiros que são maiores do que 0 são válidos.</span><span class="sxs-lookup"><span data-stu-id="7bb22-137">So this rule says that only whole numbers that are greater than 0 are valid.</span></span> 
- <span data-ttu-id="7bb22-138">O `formula1` é um número embutido.</span><span class="sxs-lookup"><span data-stu-id="7bb22-138">The `formula1` is a hard-coded number.</span></span> <span data-ttu-id="7bb22-139">Se não souber no momento da codificação qual é o valor, você também pode usar uma fórmula do Excel (como uma cadeia de caracteres) para o valor.</span><span class="sxs-lookup"><span data-stu-id="7bb22-139">If you don't know at coding time what the value should be, you can also use an Excel formula (as a string) for the value.</span></span> <span data-ttu-id="7bb22-140">Por exemplo, "= A3" e "SUM(A4,B5) =" também seriam valores `formula1`.</span><span class="sxs-lookup"><span data-stu-id="7bb22-140">For example, "=A3" and "=SUM(A4,B5)" could also be values of `formula1`.</span></span>

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

<span data-ttu-id="7bb22-141">Confira [BasicDataValidation](https://docs.microsoft.com/javascript/api/excel/excel.basicdatavalidation) para uma lista de outros operadores binários.</span><span class="sxs-lookup"><span data-stu-id="7bb22-141">See [BasicDataValidation](https://docs.microsoft.com/javascript/api/excel/excel.basicdatavalidation) for a list of the other binary operators.</span></span> 

<span data-ttu-id="7bb22-142">Também há dois ternários: "Entre" e "NotBetween".</span><span class="sxs-lookup"><span data-stu-id="7bb22-142">There are also two ternary operators: "Between" and "NotBetween".</span></span> <span data-ttu-id="7bb22-143">Para usá-los, você deve especificar a propriedade`formula2` opcional.</span><span class="sxs-lookup"><span data-stu-id="7bb22-143">To use these, you must specify the optional `formula2` property.</span></span> <span data-ttu-id="7bb22-144">Os valores`formula1` e `formula2` valores são operandos delimitadores.</span><span class="sxs-lookup"><span data-stu-id="7bb22-144">The `formula1` and `formula2` values are the bounding operands.</span></span> <span data-ttu-id="7bb22-145">O valor que o usuário tenta inserir na célula é o terceiro operando (calculado).</span><span class="sxs-lookup"><span data-stu-id="7bb22-145">The value that the user tries to enter in the cell is the third (evaluated) operand.</span></span> <span data-ttu-id="7bb22-146">Este é um exemplo de como usar o operador "Entre":</span><span class="sxs-lookup"><span data-stu-id="7bb22-146">The following is an example of using the "Between" operator:</span></span>

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

<span data-ttu-id="7bb22-147">As próximas duas regras de propriedades usam o objeto [DateTimeDataValidation](https://docs.microsoft.com/javascript/api/excel/excel.datetimedatavalidation)como o valor.</span><span class="sxs-lookup"><span data-stu-id="7bb22-147">The next two rule properties take a [DateTimeDataValidation](https://docs.microsoft.com/javascript/api/excel/excel.datetimedatavalidation) object as their value.</span></span>

- `date`
- `time`

<span data-ttu-id="7bb22-148">O`DateTimeDataValidation` objeto é estruturado da mesma forma que o `BasicDataValidation`: com as propriedades `formula1`, `formula2`, e `operator` e é usado da mesma maneira.</span><span class="sxs-lookup"><span data-stu-id="7bb22-148">The `DateTimeDataValidation` object is structured similarly to the `BasicDataValidation`: it has the properties `formula1`, `formula2`, and `operator`, and is used in the same way.</span></span> <span data-ttu-id="7bb22-149">A diferença é que você não pode usar um número nas propriedades de fórmula, mas você pode inserir uma cadeia [ISO 8606 datetime](https://www.iso.org/iso-8601-date-and-time-format.html) (ou uma fórmula do Excel).</span><span class="sxs-lookup"><span data-stu-id="7bb22-149">The difference is that you cannot use a number in the formula properties, but you can enter a [ISO 8606 datetime](https://www.iso.org/iso-8601-date-and-time-format.html) string (or an Excel formula).</span></span> <span data-ttu-id="7bb22-150">A seguir está um exemplo que define os valores válidos como datas na primeira semana de abril de 2018.</span><span class="sxs-lookup"><span data-stu-id="7bb22-150">The following is an example that defines valid values as dates in the first week of April, 2018.</span></span> 

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

### <a name="list-validation-rule-type"></a><span data-ttu-id="7bb22-151">Lista tipo de regra de validação</span><span class="sxs-lookup"><span data-stu-id="7bb22-151">List validation rule type</span></span>

<span data-ttu-id="7bb22-152">Use a `list` propriedade no `DataValidationRule` objeto para especificar valores que apenas válidos são em uma lista finita.</span><span class="sxs-lookup"><span data-stu-id="7bb22-152">Use the `list` property in the `DataValidationRule` object to specify that the only valid values are those from a finite list.</span></span> <span data-ttu-id="7bb22-153">Apresentamos um exemplo a seguir.</span><span class="sxs-lookup"><span data-stu-id="7bb22-153">The following is an example.</span></span> <span data-ttu-id="7bb22-154">Observe o seguinte sobre este código:</span><span class="sxs-lookup"><span data-stu-id="7bb22-154">Note the following about this code:</span></span>

- <span data-ttu-id="7bb22-155">Ele pressupõe que se trata de uma planilha chamada "Nomes" e que os valores no intervalo "A1: A3" são nomes.</span><span class="sxs-lookup"><span data-stu-id="7bb22-155">It assumes that there is a worksheet named "Names" and that the values in the range "A1:A3" are names.</span></span>
- <span data-ttu-id="7bb22-156">A `source` propriedade especifica a lista de valores válidos.</span><span class="sxs-lookup"><span data-stu-id="7bb22-156">The `source` property specifies the list of valid values.</span></span> <span data-ttu-id="7bb22-157">O argumento de cadeia de caracteres se refere a um intervalo que contém os nomes.</span><span class="sxs-lookup"><span data-stu-id="7bb22-157">The string argument refers to a range containing the names.</span></span> <span data-ttu-id="7bb22-158">Você também pode atribuir uma lista delimitada por vírgula; Por exemplo: "Clara, Ricky, Liz".</span><span class="sxs-lookup"><span data-stu-id="7bb22-158">You can also assign a comma-delimited list; for example: "Sue, Ricky, Liz".</span></span> 
- <span data-ttu-id="7bb22-159">A `inCellDropDown` propriedade especifica se um controle de lista suspensa será exibido na célula quando o usuário a seleciona.</span><span class="sxs-lookup"><span data-stu-id="7bb22-159">The `inCellDropDown` property specifies whether a drop-down control will appear in the cell when the user selects it.</span></span> <span data-ttu-id="7bb22-160">Se definido como `true`, em seguida, a lista suspensa é exibida com a lista de valores do `source`.</span><span class="sxs-lookup"><span data-stu-id="7bb22-160">If set to `true`, then the drop-down appears with the list of values from the `source`.</span></span>

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

### <a name="custom-validation-rule-type"></a><span data-ttu-id="7bb22-161">Tipo de regra de validação personalizada</span><span class="sxs-lookup"><span data-stu-id="7bb22-161">Custom validation rule type</span></span>

<span data-ttu-id="7bb22-162">Use a `custom` propriedade na `DataValidationRule` objeto para especificar uma fórmula de validação personalizada.</span><span class="sxs-lookup"><span data-stu-id="7bb22-162">Use the `custom` property in the `DataValidationRule` object to specify a custom validation formula.</span></span> <span data-ttu-id="7bb22-163">Apresentamos um exemplo a seguir.</span><span class="sxs-lookup"><span data-stu-id="7bb22-163">The following is an example.</span></span> <span data-ttu-id="7bb22-164">Observe o seguinte sobre este código:</span><span class="sxs-lookup"><span data-stu-id="7bb22-164">Note the following about this code:</span></span>

- <span data-ttu-id="7bb22-165">Ele pressupõe que há uma tabela de duas colunas com as colunas **Nome do Atleta** e **Comentários** nas colunas A e B da planilha.</span><span class="sxs-lookup"><span data-stu-id="7bb22-165">It assumes there is a two-column table with columns **Athlete Name** and **Comments** in the A and B columns of the worksheet.</span></span>
- <span data-ttu-id="7bb22-166">Para reduzir o nível de detalhamento na coluna **comentários**, ela torna os dados que inclui os nome do atleta inválidos.</span><span class="sxs-lookup"><span data-stu-id="7bb22-166">To reduce verbosity in the **Comments** column, it makes data that includes the athlete's name invalid.</span></span>
- <span data-ttu-id="7bb22-167">`SEARCH(A2,B2)` Retorna a posição inicial, na cadeia de caracteres em B2, da cadeia de caracteres em A2.</span><span class="sxs-lookup"><span data-stu-id="7bb22-167">`SEARCH(A2,B2)` returns the starting position, in string in B2, of the string in A2.</span></span> <span data-ttu-id="7bb22-168">Se A2 não estão contidas em B2, um número não é retornado.</span><span class="sxs-lookup"><span data-stu-id="7bb22-168">If A2 is not contained in B2, it does not return a number.</span></span> <span data-ttu-id="7bb22-169">`ISNUMBER()`retorna booliano.</span><span class="sxs-lookup"><span data-stu-id="7bb22-169">`ISNUMBER()` returns a boolean.</span></span> <span data-ttu-id="7bb22-170">Portanto a`formula` propriedade diz que os dados válidos para a coluna**Comentário** são os dados que não incluem a cadeia de caracteres da coluna **Nome do Atleta**.</span><span class="sxs-lookup"><span data-stu-id="7bb22-170">So the `formula` property says that valid data for the **Comment** column is data that does not include the string in the **Athlete Name** column.</span></span>

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

## <a name="create-validation-error-alerts"></a><span data-ttu-id="7bb22-171">Criar alertas de erro de validação</span><span class="sxs-lookup"><span data-stu-id="7bb22-171">Create validation error alerts</span></span>

<span data-ttu-id="7bb22-172">Você pode criar um alerta de erro personalizado que aparece quando um usuário tenta inserir dados inválidos em uma célula.</span><span class="sxs-lookup"><span data-stu-id="7bb22-172">You can a create custom error alert that appears when a user tries to enter invalid data in a cell.</span></span> <span data-ttu-id="7bb22-173">Apresentamos um exemplo simples a seguir.</span><span class="sxs-lookup"><span data-stu-id="7bb22-173">The following is a simple example.</span></span> <span data-ttu-id="7bb22-174">Observe o seguinte sobre este código:</span><span class="sxs-lookup"><span data-stu-id="7bb22-174">Note the following about this code:</span></span>

- <span data-ttu-id="7bb22-175">A `style` propriedade determina se o usuário obtém um alerta informativo, um aviso e um alerta "parar".</span><span class="sxs-lookup"><span data-stu-id="7bb22-175">The `style` property determines whether the user gets an informational alert, a warning, or a "stop" alert.</span></span> <span data-ttu-id="7bb22-176">Apenas `Stop` realmente impede que o usuário adicione dados inválidos.</span><span class="sxs-lookup"><span data-stu-id="7bb22-176">Only `Stop` actually prevents the user from adding invalid data.</span></span> <span data-ttu-id="7bb22-177">O pop-up para `Warning` e `Information` tem opções para permitir que o usuário insira dados inválidos assim mesmo.</span><span class="sxs-lookup"><span data-stu-id="7bb22-177">The pop-up for `Warning` and `Information` has options that allow the user enter the invalid data anyway.</span></span>
- <span data-ttu-id="7bb22-178">As `showAlert` propriedades padrão para `true`.</span><span class="sxs-lookup"><span data-stu-id="7bb22-178">The `showAlert` property defaults to `true`.</span></span> <span data-ttu-id="7bb22-179">Isso significa que o host do Excel exibirá um alerta genérico (do tipo `Stop`), a menos que você crie um alerta personalizado que defina `showAlert` para `false` ou define uma mensagem, o título e estilo personalizados.</span><span class="sxs-lookup"><span data-stu-id="7bb22-179">This means that the Excel host will pop-up a generic alert (of type `Stop`) unless you create a custom alert which either sets `showAlert` to `false` or sets a custom message, title, and style.</span></span> <span data-ttu-id="7bb22-180">O código define uma mensagem personalizada e o título.</span><span class="sxs-lookup"><span data-stu-id="7bb22-180">This code sets a custom message and title.</span></span>


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

<span data-ttu-id="7bb22-181">Para saber mais, confira [DataValidationErrorAlert](https://docs.microsoft.com/javascript/api/excel/excel.datavalidationerroralert).</span><span class="sxs-lookup"><span data-stu-id="7bb22-181">For more information, see [DataValidationErrorAlert](https://docs.microsoft.com/javascript/api/excel/excel.datavalidationerroralert).</span></span>

## <a name="create-validation-prompts"></a><span data-ttu-id="7bb22-182">Criar solicitações de validação</span><span class="sxs-lookup"><span data-stu-id="7bb22-182">Create validation prompts</span></span>

<span data-ttu-id="7bb22-183">Você pode criar um prompt instrucional que é exibido quando um usuário passa o mouse sobre ou seleciona uma célula para os dados em que foi aplicada a validação.</span><span class="sxs-lookup"><span data-stu-id="7bb22-183">You can create an instructional prompt that appears when a user hovers over, or selects, a cell to which data validation has been applied.</span></span> <span data-ttu-id="7bb22-184">Este é um exemplo:</span><span class="sxs-lookup"><span data-stu-id="7bb22-184">The following is an example:</span></span>

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

<span data-ttu-id="7bb22-185">Para saber mais, confira [DataValidationPrompt](https://docs.microsoft.com/javascript/api/excel/excel.datavalidationprompt).</span><span class="sxs-lookup"><span data-stu-id="7bb22-185">For more information, see [DataValidationPrompt](https://docs.microsoft.com/javascript/api/excel/excel.datavalidationprompt).</span></span>

## <a name="remove-data-validation-from-a-range"></a><span data-ttu-id="7bb22-186">Remover validação de dados de um intervalo</span><span class="sxs-lookup"><span data-stu-id="7bb22-186">Remove data validation from a range</span></span>

<span data-ttu-id="7bb22-187">Para remover a validação de dados de um intervalo, acionar o método [Range.dataValidation.clear()](https://docs.microsoft.com/javascript/api/excel/excel.datavalidation#clear--).</span><span class="sxs-lookup"><span data-stu-id="7bb22-187">To remove data validation from a range, call the  [Range.dataValidation.clear()](https://docs.microsoft.com/javascript/api/excel/excel.datavalidation#clear--) method.</span></span>

```js
myrange.dataValidation.clear()
```

<span data-ttu-id="7bb22-188">Não é necessário que o intervalo que você desmarcar seja o  mesmo intervalo de um intervalo no qual você adicionou a validação de dados.</span><span class="sxs-lookup"><span data-stu-id="7bb22-188">It isn't necessary that the range you clear is exactly the same range as a range on which you added data validation.</span></span> <span data-ttu-id="7bb22-189">Caso contrário, apenas as células sobrepostas, se houver, dos dois intervalos são desmarcadas.</span><span class="sxs-lookup"><span data-stu-id="7bb22-189">If it isn't, only the overlapping cells, if any, of the two ranges are cleared.</span></span> 

> [!NOTE]
> <span data-ttu-id="7bb22-190">Limpar a validação de dados de um intervalo também limpará qualquer validação de dados que o usuário tenha adicionado manualmente ao intervalo.</span><span class="sxs-lookup"><span data-stu-id="7bb22-190">Clearing data validation from a range will also clear any data validation that a user has added manually to the range.</span></span>

## <a name="see-also"></a><span data-ttu-id="7bb22-191">Confira também</span><span class="sxs-lookup"><span data-stu-id="7bb22-191">See also</span></span>

- [<span data-ttu-id="7bb22-192">Conceitos fundamentais de programação com a API JavaScript do Excel</span><span class="sxs-lookup"><span data-stu-id="7bb22-192">Fundamental programming concepts with the Excel JavaScript API</span></span>](excel-add-ins-core-concepts.md)
- [<span data-ttu-id="7bb22-193">Objeto Application (JavaScript API para Excel)</span><span class="sxs-lookup"><span data-stu-id="7bb22-193">DataValidation Object (JavaScript API for Excel)</span></span>](https://docs.microsoft.com/javascript/api/excel/excel.datavalidation)
- [<span data-ttu-id="7bb22-194">Objeto Range (API JavaScript para Excel)</span><span class="sxs-lookup"><span data-stu-id="7bb22-194">Range Object (JavaScript API for Excel)</span></span>](https://docs.microsoft.com/javascript/api/excel/excel.range)



 
