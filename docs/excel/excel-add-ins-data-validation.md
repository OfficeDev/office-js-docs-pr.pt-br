---
title: Adicionar validação de dados para intervalos do Excel
description: Saiba como as EXCEL JavaScript permitem que seu complemento adicione validação automática de dados a tabelas, colunas, linhas e outros intervalos em uma workbook.
ms.date: 03/19/2019
localization_priority: Normal
ms.openlocfilehash: 2579473800a20ba864b42b8a18b8023dff826c5e
ms.sourcegitcommit: e570fa8925204c6ca7c8aea59fbf07f73ef1a803
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 08/05/2021
ms.locfileid: "53774151"
---
# <a name="add-data-validation-to-excel-ranges"></a>Adicionar validação de dados para intervalos do Excel

A biblioteca JavaScript do Excel fornece APIs para permitir que seu suplemento adicione a validação de dados automáticos a tabelas, colunas, linhas e outros intervalos em uma pasta de trabalho. Para entender os conceitos e a terminologia da validação de dados, consulte os artigos a seguir sobre como os usuários adicionam validação de dados por meio da interface do usuário Excel usuário.

- [Aplicar validação de dados às células](https://support.office.com/article/Apply-data-validation-to-cells-29FECBCC-D1B9-42C1-9D76-EFF3CE5F7249)
- [Validação de dados](https://support.office.com/article/More-on-data-validation-f38dee73-9900-4ca6-9301-8a5f6e1f0c4c)
- [Exemplos e descrição de validação de dados no Excel](https://support.microsoft.com/help/211485/description-and-examples-of-data-validation-in-excel)

## <a name="programmatic-control-of-data-validation"></a>Controle de programação de validação de dados

A `Range.dataValidation` propriedade, que usa um objeto [DataValidation](/javascript/api/excel/excel.datavalidation), é o ponto de entrada para o controle de programação de validação de dados no Excel. Há cinco propriedades para o objeto `DataValidation`:

- `rule` &#8212;Define o que constitui dados válidos para o intervalo. Ver [DataValidationRule](/javascript/api/excel/excel.datavalidationrule).
- `errorAlert` &#8212;Especifica se um erro é exibido se o usuário insere dados inválidos e define o texto, o título e o estilo de alerta; Por exemplo, **informativo**, **Aviso**, e **Parar**. Ver [DataValidationErrorAlert](/javascript/api/excel/excel.datavalidationerroralert).
- `prompt` &#8212;Especifica se um aviso aparece quando o usuário passa o mouse sobre o intervalo e define a mensagem de aviso. Ver [DataValidationPrompt](/javascript/api/excel/excel.datavalidationprompt).
- `ignoreBlanks` &#8212;Especifica se aplica a regra de validação de dados a células em branco no intervalo. O padrão é `true`
- `type` &#8212;Identificação somente leitura do tipo de validação, como WholeNumber, data, TextLength etc. Ela é definida indiretamente quando você define a propriedade `rule`.

> [!NOTE]
> A validação de dados adicionada programaticamente funciona exatamente como a validação de dados adicionada manualmente. Em particular, observe que a validação de dados é disparada somente se o usuário inserir diretamente um valor em uma célula ou copiar e colar uma célula de outro local da pasta de trabalho e escolher a opção de colagem **Valores**. Se o usuário copiar uma célula e fizer uma colagem simples em um intervalo com a validação de dados, a validação não será disparada.

## <a name="creating-validation-rules"></a>Criar regras de validação

Para adicionar a validação de dados a um intervalo, o código deve configurar a propriedade `rule` do objeto `DataValidation` em `Range.dataValidation`. Isso leva ao objeto [DataValidationRule](/javascript/api/excel/excel.datavalidationrule) que tem sete propriedades opcionais. *Não mais de uma dessas propriedades pode estar presente em qualquer objeto `DataValidationRule`.* A propriedade que você incluir determina o tipo de validação.

### <a name="basic-and-datetime-validation-rule-types"></a>Tipos de regras de validação Basic e DateTime

As três primeiras propriedades `DataValidationRule` (ou seja, tipos de regra de validação) consideram o objeto [BasicDataValidation](/javascript/api/excel/excel.basicdatavalidation) como o seu valor.

- `wholeNumber` &#8212;Requer um número inteiro, além de outra validação especificada pelo objeto `BasicDataValidation`.
- `decimal` &#8212;Requer um número decimal, além de outra validação especificada pelo objeto `BasicDataValidation`.
- `textLength` &#8212;Aplicam-se os detalhes de validação do objeto `BasicDataValidation` ao *comprimento* de valor da célula.

Aqui está um exemplo de como criar uma regra de validação. Observe o seguinte sobre este código.

- O `operator` é o operador binário "GreaterThan". Sempre que você usa um operador binário, o valor que o usuário tenta inserir na célula é o operando à esquerda e o valor especificado em `formula1` é o operando à direita. Então esta regra diz que apenas números inteiros que são maiores do que 0 são válidos.
- O `formula1` é um número embutido. Se não souber no momento da codificação qual é o valor, você também poderá usar uma fórmula do Excel (como uma cadeia de caracteres) para o valor. Por exemplo, "= A3" e "SOMA(A4,B5) =" também seriam valores `formula1`.

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

Confira [BasicDataValidation](/javascript/api/excel/excel.basicdatavalidation) para uma lista de outros operadores binários. 

Também há dois operadores ternários: "Between" e "NotBetween". Para usá-los, você deve especificar a propriedade `formula2` opcional. Os valores`formula1` e `formula2` são os operandos delimitadores. O valor que o usuário tenta inserir na célula é o terceiro operando (calculado). A seguir, um exemplo de uso do operador "Between".

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

As próximas duas regras de propriedades usam o objeto [DateTimeDataValidation](/javascript/api/excel/excel.datetimedatavalidation) como seu valor.

- `date`
- `time`

O objeto `DateTimeDataValidation` é estruturado da mesma forma que o `BasicDataValidation`: com as propriedades `formula1`, `formula2` e `operator`, e é usado da mesma maneira. A diferença é que você não pode usar um número nas propriedades de fórmula, mas você pode inserir uma cadeia [ISO 8606 datetime](https://www.iso.org/iso-8601-date-and-time-format.html) (ou uma fórmula do Excel). A seguir está um exemplo que define os valores válidos como datas na primeira semana de abril de 2018. 

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

### <a name="list-validation-rule-type"></a>Tipos de regra de validação de lista

Use a propriedade `list` do objeto `DataValidationRule` para especificar valores que são válidos apenas em uma lista finita. Apresentamos um exemplo a seguir. Observe o seguinte sobre este código.

- Ele pressupõe que se trata de uma planilha chamada "Nomes" e que os valores no intervalo "A1: A3" são nomes.
- A propriedade `source` especifica a lista de valores válidos. O argumento de cadeia de caracteres se refere a um intervalo que contém os nomes. Você também pode atribuir uma lista delimitada por vírgula; por exemplo: "Lara, Pedro, Marina".
- A propriedade `inCellDropDown` especifica se um controle de lista suspensa será exibido na célula quando o usuário a selecionar. Se definido como `true`, em seguida, a lista suspensa é exibida com a lista de valores do `source`.

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

### <a name="custom-validation-rule-type"></a>Tipo de regra de validação personalizada

Use a propriedade `custom` no objeto `DataValidationRule` para especificar uma fórmula de validação personalizada. Apresentamos um exemplo a seguir. Observe o seguinte sobre este código.

- Ele pressupõe que há uma tabela de duas colunas com as colunas **Nome do Atleta** e **Comentários** nas colunas A e B da planilha.
- Para reduzir o nível de detalhamento na coluna **Comentários**, ela torna os dados que incluem o nome do atleta inválidos.
- `SEARCH(A2,B2)` Retorna a posição inicial, na cadeia de caracteres em B2, da cadeia de caracteres em A2. Se A2 não estiver contida em B2, ela não retornará um número. `ISNUMBER()` retorna um booliano. Portanto, a propriedade `formula` diz que os dados válidos para a coluna **Comentário** são os dados que não incluem a cadeia de caracteres da coluna **Nome do Atleta**.

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

## <a name="create-validation-error-alerts"></a>Criar alertas de erro de validação

Você pode criar um alerta de erro personalizado que aparece quando um usuário tenta inserir dados inválidos em uma célula. Apresentamos um exemplo simples a seguir. Observe o seguinte sobre este código.

- A propriedade `style` determina se o usuário recebe um alerta informativo, um aviso e um alerta "parar". Apenas `Stop` realmente impede que o usuário adicione dados inválidos. O pop-up para `Warning` e `Information` tem opções para permitir que o usuário insira dados inválidos assim mesmo.
- As propriedades `showAlert` padrão para `true`. Isso significa Excel um alerta genérico (de tipo), a menos que você crie um alerta personalizado que define ou define uma mensagem, título e `Stop` `showAlert` estilo `false` personalizados. O código define uma mensagem personalizada e o título.

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

Para saber mais, confira [DataValidationErrorAlert](/javascript/api/excel/excel.datavalidationerroralert).

## <a name="create-validation-prompts"></a>Criar solicitações de validação

Você pode criar um prompt instrutivo que é exibido quando um usuário passa o mouse sobre ele ou seleciona uma célula à qual os dados de validação foram aplicados. Apresentamos um exemplo a seguir.

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

Para saber mais, confira [DataValidationPrompt](/javascript/api/excel/excel.datavalidationprompt).

## <a name="remove-data-validation-from-a-range"></a>Remover validação de dados de um intervalo

Para remover a validação de dados de um intervalo, acione o método [Range.dataValidation.clear()](/javascript/api/excel/excel.datavalidation#clear__).

```js
myrange.dataValidation.clear()
```

Não é necessário que o intervalo que você desmarcar seja o mesmo intervalo de um intervalo no qual você adicionou a validação de dados. Caso contrário, apenas as células sobrepostas, se houver, dos dois intervalos são desmarcadas. 

> [!NOTE]
> Limpar a validação de dados de um intervalo também limpará qualquer validação de dados que o usuário tenha adicionado manualmente ao intervalo.

## <a name="see-also"></a>Confira também

- [Modelo de objeto JavaScript do Excel em Suplementos do Office](excel-add-ins-core-concepts.md)
- [Objeto Application (JavaScript API para Excel)](/javascript/api/excel/excel.datavalidation)
- [Objeto Range (API JavaScript para Excel)](/javascript/api/excel/excel.range)
