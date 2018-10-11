---
title: Adicionar validação de dados a intervalos do Excel
description: ''
ms.date: 10/03/2018
ms.openlocfilehash: 9e3aba8d87e84405bb3e1ae35a8d35d60ce8e2b6
ms.sourcegitcommit: 563c53bac52b31277ab935f30af648f17c5ed1e2
ms.translationtype: HT
ms.contentlocale: pt-BR
ms.lasthandoff: 10/10/2018
ms.locfileid: "25459151"
---
# <a name="add-data-validation-to-excel-ranges"></a><span data-ttu-id="760b2-102">Adicionar validação de dados a intervalos do Excel</span><span class="sxs-lookup"><span data-stu-id="760b2-102">Add data validation to Excel ranges</span></span>

<span data-ttu-id="760b2-p101">Biblioteca JavaScript do Excel fornece APIs para habilitar o suplemento para adicionar a validação de dados automática a tabelas, linhas, colunas e outros intervalos em uma pasta de trabalho. Para entender os conceitos e a terminologia de validação de dados, consulte os seguintes artigos sobre como os usuários adicionam validação de dados por meio da interface de usuário do Excel:</span><span class="sxs-lookup"><span data-stu-id="760b2-p101">The Excel JavaScript Library provides APIs to enable your add-in to add automatic data validation to tables, columns, rows, and other ranges in a workbook. To understand the concepts and the terminology of data validation, please see the following articles about how users add data validation through the Excel UI:</span></span>

- [<span data-ttu-id="760b2-105">Aplicar validação de dados a células</span><span class="sxs-lookup"><span data-stu-id="760b2-105">Apply data validation to cells</span></span>](https://support.office.com/article/Apply-data-validation-to-cells-29FECBCC-D1B9-42C1-9D76-EFF3CE5F7249)
- [<span data-ttu-id="760b2-106">Mais sobre validação de dados</span><span class="sxs-lookup"><span data-stu-id="760b2-106">More on data validation</span></span>](https://support.office.com/article/More-on-data-validation-f38dee73-9900-4ca6-9301-8a5f6e1f0c4c)
- [<span data-ttu-id="760b2-107">Descrição e exemplos de validação de dados no Excel</span><span class="sxs-lookup"><span data-stu-id="760b2-107">Description and examples of data validation in Excel</span></span>](https://support.microsoft.com/help/211485/description-and-examples-of-data-validation-in-excel)

## <a name="programmatic-control-of-data-validation"></a><span data-ttu-id="760b2-108">Controle programático de validação de dados</span><span class="sxs-lookup"><span data-stu-id="760b2-108">Programmatic control of data validation</span></span>

<span data-ttu-id="760b2-p102">A propriedade `Range.dataValidation`, que usa um objeto [DataValidation](https://docs.microsoft.com/javascript/api/excel/excel.datavalidation) , é o ponto de entrada para o controle programático da validação de dados no Excel. Há cinco propriedades para o objeto `DataValidation`:</span><span class="sxs-lookup"><span data-stu-id="760b2-p102">The `Range.dataValidation` property, which takes a [DataValidation](https://docs.microsoft.com/javascript/api/excel/excel.datavalidation) object, is the entry point for programmatic control of data validation in Excel. There are five properties to the `DataValidation` object:</span></span>

- <span data-ttu-id="760b2-p103">`rule` – Define o que constitui dados válidos para o intervalo. Confira [DataValidationRule](https://docs.microsoft.com/javascript/api/excel/excel.datavalidationrule).</span><span class="sxs-lookup"><span data-stu-id="760b2-p103">`rule` &#8212; Defines what constitutes valid data for the range. See [DataValidationRule](https://docs.microsoft.com/javascript/api/excel/excel.datavalidationrule).</span></span>
- <span data-ttu-id="760b2-p104">`errorAlert` – Especifica se um erro será exibido caso o usuário insira dados inválidos, e define o texto, o título e o estilo do alerta, por exemplo: **Informativo**, **Aviso**e **Parar**. Confira [DataValidationErrorAlert](https://docs.microsoft.com/javascript/api/excel/excel.datavalidationerroralert).</span><span class="sxs-lookup"><span data-stu-id="760b2-p104">`errorAlert` &#8212; Specifies whether an error pops up if the user enters invalid data, and defines the alert text, title, and style; for example, **Informational**, **Warning**, and **Stop**. See [DataValidationErrorAlert](https://docs.microsoft.com/javascript/api/excel/excel.datavalidationerroralert).</span></span>
- <span data-ttu-id="760b2-p105">`prompt` – Especifica se uma solicitação será exibida quando o usuário passar o mouse sobre o intervalo e define a mensagem da solicitação. Confira [DataValidationPrompt](https://docs.microsoft.com/javascript/api/excel/excel.datavalidationprompt).</span><span class="sxs-lookup"><span data-stu-id="760b2-p105">`prompt` &#8212; Specifies whether a prompt appears when the user hovers over the range and defines the prompt message. See [DataValidationPrompt](https://docs.microsoft.com/javascript/api/excel/excel.datavalidationprompt).</span></span>
- <span data-ttu-id="760b2-p106">`ignoreBlanks` – Especifica se a regra de validação de dados se aplica a células em branco no intervalo. Padrão para `true`.</span><span class="sxs-lookup"><span data-stu-id="760b2-p106">`ignoreBlanks` &#8212; Specifies whether the data validation rule applies to blank cells in the range. Defaults to `true`.</span></span>
- <span data-ttu-id="760b2-119">`type` – Uma identificação somente leitura do tipo de validação, como WholeNumber, Date, TextLength etc. Ela é definida indiretamente ao se definir a propriedade `rule`.</span><span class="sxs-lookup"><span data-stu-id="760b2-119">`type` &#8212; A read-only identification of the validation type, such as WholeNumber, Date, TextLength, etc. It is set indirectly when you set the `rule` property.</span></span>

> [!NOTE]
> <span data-ttu-id="760b2-p107">A validação de dados adicionada de forma programática se comporta exatamente como manualmente adicionada a validação de dados. Em particular, observe que a validação de dados é acionada apenas se o usuário insere um valor em uma célula ou copia e cola uma célula de qualquer outro lugar na pasta de trabalho e escolhe  a opção de colagem **Valores**. Se o usuário copiar uma célula e fizer uma colagem sem formatação em um intervalo com validação de dados, a validação não será acionada.</span><span class="sxs-lookup"><span data-stu-id="760b2-p107">Data validation added programmatically behaves just like manually added data validation. In particular, note that data validation is triggered only if the user directly enters a value into a cell or copies and pastes a cell from elsewhere in the workbook and chooses the **Values** paste option. If the user copies a cell and does a plain paste into a range with data validation, validation is not triggered.</span></span>

## <a name="creating-validation-rules"></a><span data-ttu-id="760b2-123">Criando regras de validação</span><span class="sxs-lookup"><span data-stu-id="760b2-123">Creating validation rules</span></span>

<span data-ttu-id="760b2-p108">Para adicionar a validação de dados a um intervalo, seu código deve definir a propriedade `rule` do objeto `DataValidation` em `Range.dataValidation`. Isso leva a um objeto [DataValidationRule](https://docs.microsoft.com/javascript/api/excel/excel.datavalidationrule) que tem sete propriedades opcionais. *Não mais de uma dessas propriedades pode estar presente em qualquer `DataValidationRule` objeto.* A propriedade que você incluir determina o tipo de validação.</span><span class="sxs-lookup"><span data-stu-id="760b2-p108">To add data validation to a range, your code must set the `rule` property of the `DataValidation` object in `Range.dataValidation`. This takes a [DataValidationRule](https://docs.microsoft.com/javascript/api/excel/excel.datavalidationrule) object which has seven optional properties. *No more than one of these properties may be present in any `DataValidationRule` object.* The property that you include determines the type of validation.</span></span>

### <a name="basic-and-datetime-validation-rule-types"></a><span data-ttu-id="760b2-128">Tipos de regra de validação Basic e DateTime</span><span class="sxs-lookup"><span data-stu-id="760b2-128">Basic and DateTime validation rule types</span></span>

<span data-ttu-id="760b2-129">As três primeiras propriedades `DataValidationRule` (isto é, tipos de regra de validação) usam um objeto [BasicDataValidation](https://docs.microsoft.com/javascript/api/excel/excel.basicdatavalidation) como seu valor.</span><span class="sxs-lookup"><span data-stu-id="760b2-129">The first three `DataValidationRule` properties (i.e., validation rule types) take a [BasicDataValidation](https://docs.microsoft.com/javascript/api/excel/excel.basicdatavalidation) object as their value.</span></span>

- <span data-ttu-id="760b2-130">`wholeNumber` – Requer um número inteiro, além de qualquer outra validação especificada pelo objeto `BasicDataValidation`.</span><span class="sxs-lookup"><span data-stu-id="760b2-130">`wholeNumber` &#8212; Requires a whole number in addition to any other validation specified by the `BasicDataValidation` object.</span></span>
- <span data-ttu-id="760b2-131">`decimal` – Requer um número decimal, além de qualquer outra validação especificada pelo objeto `BasicDataValidation`.</span><span class="sxs-lookup"><span data-stu-id="760b2-131">`decimal` &#8212; Requires a decimal number in addition to any other validation specified by the `BasicDataValidation` object.</span></span>
- <span data-ttu-id="760b2-132">`textLength` – Aplica os detalhes de validação no objeto `BasicDataValidation` ao *comprimento* do valor da célula.</span><span class="sxs-lookup"><span data-stu-id="760b2-132">`textLength` &#8212; Applies the validation details in the `BasicDataValidation` object to the *length* of the cell's value.</span></span>

<span data-ttu-id="760b2-p109">Este é um exemplo de como criar uma regra de validação. Observe o seguinte sobre este código:</span><span class="sxs-lookup"><span data-stu-id="760b2-p109">Here is an example of creating a validation rule. Note the following about this code:</span></span>

- <span data-ttu-id="760b2-p110">O `operator` é o operador binário "GreaterThan". Sempre que você usar um operador binário, o valor que o usuário tentar inserir na célula é o operando esquerdo e o valor especificado em `formula1` é o operando direito. Portanto, esta regra diz que apenas os números inteiros maiores que 0 são válidos.</span><span class="sxs-lookup"><span data-stu-id="760b2-p110">The `operator` is the binary operator "GreaterThan". Whenever you use a binary operator, the value that the user tries to enter in the cell is the left-hand operand and the value specified in `formula1` is the right-hand operand. So this rule says that only whole numbers that are greater than 0 are valid.</span></span> 
- <span data-ttu-id="760b2-p111">O `formula1` é um número codificado. Se não souber o que valor deve ser no momento da codificação, você também pode usar uma fórmula do Excel (como uma sequência de caracteres) para o valor. Por exemplo, "= A3" e "= SUM(A4,B5)" também poderiam ser valores de `formula1`.</span><span class="sxs-lookup"><span data-stu-id="760b2-p111">The `formula1` is a hard-coded number. If you don't know at coding time what the value should be, you can also use an Excel formula (as a string) for the value. For example, "=A3" and "=SUM(A4,B5)" could also be values of `formula1`.</span></span>

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

<span data-ttu-id="760b2-141">Consulte [BasicDataValidation](https://docs.microsoft.com/javascript/api/excel/excel.basicdatavalidation) para obter uma lista dos outros operadores binários.</span><span class="sxs-lookup"><span data-stu-id="760b2-141">See [BasicDataValidation](https://docs.microsoft.com/javascript/api/excel/excel.basicdatavalidation) for a list of the other binary operators.</span></span> 

<span data-ttu-id="760b2-p112">Também há dois operadores ternários: "Between" e "NotBetween". Para usá-los, é preciso especificar a propriedade opcional `formula2`. Os valores `formula1` e `formula2` são os operandos delimitadores. O valor que o usuário tenta inserir na célula é o terceiro operando (avaliado). Este é um exemplo de utilização do operador "Between":</span><span class="sxs-lookup"><span data-stu-id="760b2-p112">There are also two ternary operators: "Between" and "NotBetween". To use these, you must specify the optional `formula2` property. The `formula1` and `formula2` values are the bounding operands. The value that the user tries to enter in the cell is the third (evaluated) operand. The following is an example of using the "Between" operator:</span></span>

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

<span data-ttu-id="760b2-147">As próximas duas propriedades da regra usam o objeto [DateTimeDataValidation](https://docs.microsoft.com/javascript/api/excel/excel.datetimedatavalidation) como seu valor.</span><span class="sxs-lookup"><span data-stu-id="760b2-147">The next two rule properties take a [DateTimeDataValidation](https://docs.microsoft.com/javascript/api/excel/excel.datetimedatavalidation) object as their value.</span></span>

- `date`
- `time`

<span data-ttu-id="760b2-p113">O objeto `DateTimeDataValidation` é estruturado da mesma forma que o `BasicDataValidation`: ele tem as propriedades `formula1`, `formula2` e `operator` e é usado da mesma maneira. A diferença é que você não pode usar um número nas propriedades da fórmulas, mas pode inserir uma sequência de caracteres [ISO 8606 datetime](https://www.iso.org/iso-8601-date-and-time-format.html) (ou uma fórmula do Excel). Este é um exemplo que define os valores válidos como datas na primeira semana de abril de 2018.</span><span class="sxs-lookup"><span data-stu-id="760b2-p113">The `DateTimeDataValidation` object is structured similarly to the `BasicDataValidation`: it has the properties `formula1`, `formula2`, and `operator`, and is used in the same way. The difference is that you cannot use a number in the formula properties, but you can enter a [ISO 8606 datetime](https://www.iso.org/iso-8601-date-and-time-format.html) string (or an Excel formula). The following is an example that defines valid values as dates in the first week of April, 2018.</span></span> 

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

### <a name="list-validation-rule-type"></a><span data-ttu-id="760b2-151">Tipo de regra de validação de lista</span><span class="sxs-lookup"><span data-stu-id="760b2-151">List validation rule type</span></span>

<span data-ttu-id="760b2-p114">Use a propriedade `list` no objeto `DataValidationRule` para especificar que os únicos valores válidos são aqueles de uma lista finita. Este é um exemplo. Observe o seguinte sobre este código:</span><span class="sxs-lookup"><span data-stu-id="760b2-p114">Use the `list` property in the `DataValidationRule` object to specify that the only valid values are those from a finite list. The following is an example. Note the following about this code:</span></span>

- <span data-ttu-id="760b2-155">Ele pressupõe que há uma planilha chamada “Nomes” e que os valores no intervalo “A1:A3” são nomes.</span><span class="sxs-lookup"><span data-stu-id="760b2-155">It assumes that there is a worksheet named "Names" and that the values in the range "A1:A3" are names.</span></span>
- <span data-ttu-id="760b2-p115">A propriedade `source` especifica a lista de valores válidos. O intervalo com os nomes foi atribuído a ela. Você também pode atribuir uma lista delimitada por vírgula; por exemplo: "Sue, Ricky, Liz".</span><span class="sxs-lookup"><span data-stu-id="760b2-p115">The `source` property specifies the list of valid values. The range with the names has been assigned to it. You can also assign a comma-delimited list; for example: "Sue, Ricky, Liz".</span></span> 
- <span data-ttu-id="760b2-p116">A propriedade `inCellDropDown` especifica se um controle da lista suspensa será exibido na célula quando o usuário o selecionar. Se for definido como `true`, a lista suspensa será exibida com a lista de valores de `source`.</span><span class="sxs-lookup"><span data-stu-id="760b2-p116">The `inCellDropDown` property specifies whether a drop-down control will appear in the cell when the user selects it. If set to `true`, then the drop-down appears with the list of values from the `source`.</span></span>

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

### <a name="custom-validation-rule-type"></a><span data-ttu-id="760b2-161">Tipo de regra de validação personalizada</span><span class="sxs-lookup"><span data-stu-id="760b2-161">Custom validation rule type</span></span>

<span data-ttu-id="760b2-p117">Use a propriedade `custom` no objeto `DataValidationRule` para especificar uma fórmula de validação personalizada. Este é um exemplo. Observe o seguinte sobre este código:</span><span class="sxs-lookup"><span data-stu-id="760b2-p117">Use the `custom` property in the `DataValidationRule` object to specify a custom validation formula. The following is an example. Note the following about this code:</span></span>

- <span data-ttu-id="760b2-165">Ele pressupõe que há uma tabela de duas colunas com as colunas **Nome do Atleta** e **Comentários** nas colunas A e B da planilha.</span><span class="sxs-lookup"><span data-stu-id="760b2-165">It assumes there is a two-column table with columns **Athlete Name** and **Comments** in the A and B columns of the worksheet.</span></span>
- <span data-ttu-id="760b2-166">Para reduzir o detalhamento na coluna **Comentários**, ele faz com que os dados que incluem o nome do atleta se tornem inválidos.</span><span class="sxs-lookup"><span data-stu-id="760b2-166">To reduce verbosity in the **Comments** column, it makes data that includes the athlete's name invalid.</span></span>
- <span data-ttu-id="760b2-p118">`SEARCH(A2,B2)` retorna a posição inicial, na sequência de caracteres em B2, da sequência de caracteres em A2. Se A2 não estiver contido em B2, ele não retorna um número. `ISNUMBER()` retorna um booleano. Portanto, a propriedade `formula` diz que os dados válidos para a coluna **Comentários** são dados que não incluem a sequência de caracteres na coluna **Nome do atleta** .</span><span class="sxs-lookup"><span data-stu-id="760b2-p118">`SEARCH(A2,B2)` returns the starting position, in string in B2, of the string in A2. If A2 is not contained in B2, it does not return a number. `ISNUMBER()` returns a boolean. So the `formula` property says that valid data for the **Comment** column is data that does not include the string in the **Athlete Name** column.</span></span>

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

## <a name="create-validation-error-alerts"></a><span data-ttu-id="760b2-171">Criar alertas de erro de validação</span><span class="sxs-lookup"><span data-stu-id="760b2-171">Create validation error alerts</span></span>

<span data-ttu-id="760b2-p119">Você pode criar um alerta de erro personalizado que aparece quando um usuário tentar inserir dados inválidos em uma célula. Este é um exemplo simples. Observe o seguinte sobre este código:</span><span class="sxs-lookup"><span data-stu-id="760b2-p119">You can a create custom error alert that appears when a user tries to enter invalid data in a cell. The following is a simple example. Note the following about this code:</span></span>

- <span data-ttu-id="760b2-p120">A propriedade `style` determina se o usuário obtém um alerta informativo, um aviso ou um alerta de "parar". Somente `Stop` pode realmente impedir que o usuário adicione dados inválidos. A janela pop-up para `Warning` e `Information` tem opções que permitem que o usuário insira dados inválidos mesmo assim.</span><span class="sxs-lookup"><span data-stu-id="760b2-p120">The `style` property determines whether the user gets an informational alert, a warning, or a "stop" alert. Only `Stop` actually prevents the user from adding invalid data. The pop-up for `Warning` and `Information` has options that allow the user enter the invalid data anyway.</span></span>
- <span data-ttu-id="760b2-p121">A propriedade `showAlert` assume `true` como padrão. Isso significa que o host do Excel irá abrir um alerta genérico (do tipo `Stop`), a menos que você crie um alerta personalizado que define `showAlert` para `false` ou define uma mensagem personalizada, título e estilo. Este código define uma mensagem personalizada e um título.</span><span class="sxs-lookup"><span data-stu-id="760b2-p121">The `showAlert` property defaults to `true`. This means that the Excel host will pop-up a generic alert (of type `Stop`) unless you create a custom alert which either sets `showAlert` to `false` or sets a custom message, title, and style. This code sets a custom message and title.</span></span>


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

<span data-ttu-id="760b2-181">Para obter mais informações, confira [DataValidationErrorAlert](https://docs.microsoft.com/javascript/api/excel/excel.datavalidationerroralert).</span><span class="sxs-lookup"><span data-stu-id="760b2-181">For more information, see [](https://docs.microsoft.com/javascript/api/excel/excel.datavalidationerroralert).</span></span>

## <a name="create-validation-prompts"></a><span data-ttu-id="760b2-182">Criar solicitações de validação</span><span class="sxs-lookup"><span data-stu-id="760b2-182">Create validation prompts</span></span>

<span data-ttu-id="760b2-p122">É possível criar uma solicitação de instrução que aparece quando um usuário seleciona ou passa o mouse sobre uma célula na qual a validação de dados foi aplicada. Este é um exemplo:</span><span class="sxs-lookup"><span data-stu-id="760b2-p122">You can create an instructional prompt that appears when a user hovers over, or selects, a cell to which data validation has been applied. The following is an example:</span></span>

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

<span data-ttu-id="760b2-185">Para obter mais informações, confira [DataValidationPrompt](https://docs.microsoft.com/javascript/api/excel/excel.datavalidationprompt).</span><span class="sxs-lookup"><span data-stu-id="760b2-185">For more information, see [](https://docs.microsoft.com/javascript/api/excel/excel.datavalidationprompt).</span></span>

## <a name="remove-data-validation-from-a-range"></a><span data-ttu-id="760b2-186">Remover a validação de dados de um intervalo</span><span class="sxs-lookup"><span data-stu-id="760b2-186">Remove data validation from a range</span></span>

<span data-ttu-id="760b2-187">Para remover a validação de dados de um intervalo, chame o método [Range.dataValidation.clear()](https://docs.microsoft.com/javascript/api/excel/excel.datavalidation#clear--).</span><span class="sxs-lookup"><span data-stu-id="760b2-187">To remove data validation from a range, call the  [Range.dataValidation.clear()](https://docs.microsoft.com/javascript/api/excel/excel.datavalidation#clear--) method.</span></span>

```js
myrange.dataValidation.clear()
```

<span data-ttu-id="760b2-p123">Não é necessário que o intervalo que você desmarcar seja exatamente o mesmo intervalo no qual você adicionou a validação de dados. Se não for, somente as células sobrepostas, se houver, de dois intervalos serão desmarcadas.</span><span class="sxs-lookup"><span data-stu-id="760b2-p123">It isn't necessary that the range you clear is exactly the same range as a range on which you added data validation. If it isn't, only the overlapping cells, if any, of the two ranges are cleared.</span></span> 

> [!NOTE]
> <span data-ttu-id="760b2-190">A desmarcação da validação de dados de um intervalo também será aplicada em qualquer validação de dados que um usuário tiver adicionado manualmente ao intervalo.</span><span class="sxs-lookup"><span data-stu-id="760b2-190">Clearing data validation from a range will also clear any data validation that a user has added manually to the range.</span></span>

## <a name="see-also"></a><span data-ttu-id="760b2-191">Confira também</span><span class="sxs-lookup"><span data-stu-id="760b2-191">See also</span></span>

- [<span data-ttu-id="760b2-192">Conceitos fundamentais de programação com a API JavaScript do Excel</span><span class="sxs-lookup"><span data-stu-id="760b2-192">Fundamental programming concepts with the Excel JavaScript API</span></span>](excel-add-ins-core-concepts.md)
- [<span data-ttu-id="760b2-193">Objeto DataValidation (API JavaScript para Excel)</span><span class="sxs-lookup"><span data-stu-id="760b2-193">Chart Object (JavaScript API for Excel)</span></span>](https://docs.microsoft.com/javascript/api/excel/excel.datavalidation)
- [<span data-ttu-id="760b2-194">Objeto Range (API JavaScript para Excel)</span><span class="sxs-lookup"><span data-stu-id="760b2-194">Range Object (JavaScript API for Excel)</span></span>](https://docs.microsoft.com/javascript/api/excel/excel.range)



 
