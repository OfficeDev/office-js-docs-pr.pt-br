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
# <a name="add-data-validation-to-excel-ranges"></a>Adicionar validação de dados para intervalos do Excel

A biblioteca JavaScript do Excel fornece APIs para permitir que seu suplemento adicione a validação de dados automáticos a tabelas, colunas, linhas e outros intervalos em uma pasta de trabalho. Para entender os conceitos e a terminologia de validação de dados, confira os seguintes artigos sobre como os usuários adicionam a validação de dados na interface do usuário do Excel:

- [Apply data validation to cells](https://support.office.com/article/Apply-data-validation-to-cells-29FECBCC-D1B9-42C1-9D76-EFF3CE5F7249)
- [Validação de dados](https://support.office.com/article/More-on-data-validation-f38dee73-9900-4ca6-9301-8a5f6e1f0c4c)
- [Exemplos e descrição de validação de dados no Excel](https://support.microsoft.com/help/211485/description-and-examples-of-data-validation-in-excel)

## <a name="programmatic-control-of-data-validation"></a>Controle de programação de validação de dados

A `Range.dataValidation` propriedade, que usa um objeto [DataValidation](https://docs.microsoft.com/javascript/api/excel/excel.datavalidation), é o ponto de entrada para o controle de programação de validação de dados no Excel. Há cinco propriedades a `DataValidation` objeto:

- `rule` &#8212;Define o que constitui dados válidos para o intervalo. Ver [DataValidationRule](https://docs.microsoft.com/javascript/api/excel/excel.datavalidationrule).
- `errorAlert` &#8212;Especifica se um erro é exibido se o usuário insere dados inválidos e define o texto, o título e o estilo de alerta; Por exemplo, **informativo**, **Aviso**, e **Parar**. Ver [DataValidationErrorAlert](https://docs.microsoft.com/javascript/api/excel/excel.datavalidationerroralert).
- `prompt` &#8212;Especifica se um aviso aparece quando o usuário passa o mouse sobre o intervalo e define a mensagem de aviso. Ver [DataValidationPrompt](https://docs.microsoft.com/javascript/api/excel/excel.datavalidationprompt).
- `ignoreBlanks` &#8212;Especifica se aplica a regra de validação de dados a células em branco no intervalo. O padrão é `true`
- `type` &#8212;Identificação de somente leitura do tipo de validação, como WholeNumber, data, TextLength, etc. Ela é definida indiretamente quando você define a `rule` propriedade.

> [!NOTE]
> A validação de dados adicionada programaticamente funciona exatamente como a validação de dados adicionada manualmente. Em particular, observe que a validação de dados é disparada somente se o usuário inserir diretamente um valor em uma célula ou copiar e colar uma célula de outro local da pasta de trabalho e escolher os **valores** opção de colagem. Se o usuário copiar uma célula e fazer uma colagem simples em um intervalo com a validação de dados, a validação não é disparada.

## <a name="creating-validation-rules"></a>Criar regras de validação

Para adicionar a validação de dados em um intervalo, o código deve configurar a`rule` propriedade do `DataValidation` objeto em `Range.dataValidation`. Isso leva ao objeto [DataValidationRule](https://docs.microsoft.com/javascript/api/excel/excel.datavalidationrule) que tem sete propriedades opcionais. *Não mais de uma dessas propriedades pode estar presente em qualquer `DataValidationRule` objeto.* A propriedade que você incluir determina o tipo de validação.

### <a name="basic-and-datetime-validation-rule-types"></a>Tipos de regras de validação do Basic and DateTime

As três primeiras `DataValidationRule` propriedades (ou seja, tipos de regra de validação) considere o objeto [BasicDataValidation](https://docs.microsoft.com/javascript/api/excel/excel.basicdatavalidation), como o valor.

- `wholeNumber` &#8212;Requer um número inteiro, além de outra validação especificado pelo `BasicDataValidation` objeto.
- `decimal` &#8212;Requer um número decimal, além de outra validação especificada pelo `BasicDataValidation` objeto.
- `textLength` &#8212;Aplica-se os detalhes de validação no `BasicDataValidation` objeto para o *comprimento* de valor da célula.

Aqui está um exemplo de como criar uma regra de validação. Observe o seguinte sobre este código:

- O `operator` é o operador binário "GreaterThan". Sempre que você usa um operador binário, o valor que o usuário tenta inserir na célula é operado à esquerda e o valor especificado no `formula1` é operado à direita. Então esta regra diz que apenas números inteiros que são maiores do que 0 são válidos. 
- O `formula1` é um número embutido. Se não souber no momento da codificação qual é o valor, você também pode usar uma fórmula do Excel (como uma cadeia de caracteres) para o valor. Por exemplo, "= A3" e "SUM(A4,B5) =" também seriam valores `formula1`.

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

Confira [BasicDataValidation](https://docs.microsoft.com/javascript/api/excel/excel.basicdatavalidation) para uma lista de outros operadores binários. 

Também há dois ternários: "Entre" e "NotBetween". Para usá-los, você deve especificar a propriedade`formula2` opcional. Os valores`formula1` e `formula2` valores são operandos delimitadores. O valor que o usuário tenta inserir na célula é o terceiro operando (calculado). Este é um exemplo de como usar o operador "Entre":

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

As próximas duas regras de propriedades usam o objeto [DateTimeDataValidation](https://docs.microsoft.com/javascript/api/excel/excel.datetimedatavalidation)como o valor.

- `date`
- `time`

O`DateTimeDataValidation` objeto é estruturado da mesma forma que o `BasicDataValidation`: com as propriedades `formula1`, `formula2`, e `operator` e é usado da mesma maneira. A diferença é que você não pode usar um número nas propriedades de fórmula, mas você pode inserir uma cadeia [ISO 8606 datetime](https://www.iso.org/iso-8601-date-and-time-format.html) (ou uma fórmula do Excel). A seguir está um exemplo que define os valores válidos como datas na primeira semana de abril de 2018. 

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

### <a name="list-validation-rule-type"></a>Lista tipo de regra de validação

Use a `list` propriedade no `DataValidationRule` objeto para especificar valores que apenas válidos são em uma lista finita. Apresentamos um exemplo a seguir. Observe o seguinte sobre este código:

- Ele pressupõe que se trata de uma planilha chamada "Nomes" e que os valores no intervalo "A1: A3" são nomes.
- A `source` propriedade especifica a lista de valores válidos. O argumento de cadeia de caracteres se refere a um intervalo que contém os nomes. Você também pode atribuir uma lista delimitada por vírgula; Por exemplo: "Clara, Ricky, Liz". 
- A `inCellDropDown` propriedade especifica se um controle de lista suspensa será exibido na célula quando o usuário a seleciona. Se definido como `true`, em seguida, a lista suspensa é exibida com a lista de valores do `source`.

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

Use a `custom` propriedade na `DataValidationRule` objeto para especificar uma fórmula de validação personalizada. Apresentamos um exemplo a seguir. Observe o seguinte sobre este código:

- Ele pressupõe que há uma tabela de duas colunas com as colunas **Nome do Atleta** e **Comentários** nas colunas A e B da planilha.
- Para reduzir o nível de detalhamento na coluna **comentários**, ela torna os dados que inclui os nome do atleta inválidos.
- `SEARCH(A2,B2)` Retorna a posição inicial, na cadeia de caracteres em B2, da cadeia de caracteres em A2. Se A2 não estão contidas em B2, um número não é retornado. `ISNUMBER()`retorna booliano. Portanto a`formula` propriedade diz que os dados válidos para a coluna**Comentário** são os dados que não incluem a cadeia de caracteres da coluna **Nome do Atleta**.

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

Você pode criar um alerta de erro personalizado que aparece quando um usuário tenta inserir dados inválidos em uma célula. Apresentamos um exemplo simples a seguir. Observe o seguinte sobre este código:

- A `style` propriedade determina se o usuário obtém um alerta informativo, um aviso e um alerta "parar". Apenas `Stop` realmente impede que o usuário adicione dados inválidos. O pop-up para `Warning` e `Information` tem opções para permitir que o usuário insira dados inválidos assim mesmo.
- As `showAlert` propriedades padrão para `true`. Isso significa que o host do Excel exibirá um alerta genérico (do tipo `Stop`), a menos que você crie um alerta personalizado que defina `showAlert` para `false` ou define uma mensagem, o título e estilo personalizados. O código define uma mensagem personalizada e o título.


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

Para saber mais, confira [DataValidationErrorAlert](https://docs.microsoft.com/javascript/api/excel/excel.datavalidationerroralert).

## <a name="create-validation-prompts"></a>Criar solicitações de validação

Você pode criar um prompt instrucional que é exibido quando um usuário passa o mouse sobre ou seleciona uma célula para os dados em que foi aplicada a validação. Este é um exemplo:

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

Para saber mais, confira [DataValidationPrompt](https://docs.microsoft.com/javascript/api/excel/excel.datavalidationprompt).

## <a name="remove-data-validation-from-a-range"></a>Remover validação de dados de um intervalo

Para remover a validação de dados de um intervalo, acionar o método [Range.dataValidation.clear()](https://docs.microsoft.com/javascript/api/excel/excel.datavalidation#clear--).

```js
myrange.dataValidation.clear()
```

Não é necessário que o intervalo que você desmarcar seja o  mesmo intervalo de um intervalo no qual você adicionou a validação de dados. Caso contrário, apenas as células sobrepostas, se houver, dos dois intervalos são desmarcadas. 

> [!NOTE]
> Limpar a validação de dados de um intervalo também limpará qualquer validação de dados que o usuário tenha adicionado manualmente ao intervalo.

## <a name="see-also"></a>Confira também

- [Conceitos fundamentais de programação com a API JavaScript do Excel](excel-add-ins-core-concepts.md)
- [Objeto Application (JavaScript API para Excel)](https://docs.microsoft.com/javascript/api/excel/excel.datavalidation)
- [Objeto Range (API JavaScript para Excel)](https://docs.microsoft.com/javascript/api/excel/excel.range)



 
