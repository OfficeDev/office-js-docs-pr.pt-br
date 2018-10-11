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
# <a name="add-data-validation-to-excel-ranges"></a>Adicionar validação de dados a intervalos do Excel

Biblioteca JavaScript do Excel fornece APIs para habilitar o suplemento para adicionar a validação de dados automática a tabelas, linhas, colunas e outros intervalos em uma pasta de trabalho. Para entender os conceitos e a terminologia de validação de dados, consulte os seguintes artigos sobre como os usuários adicionam validação de dados por meio da interface de usuário do Excel:

- [Aplicar validação de dados a células](https://support.office.com/article/Apply-data-validation-to-cells-29FECBCC-D1B9-42C1-9D76-EFF3CE5F7249)
- [Mais sobre validação de dados](https://support.office.com/article/More-on-data-validation-f38dee73-9900-4ca6-9301-8a5f6e1f0c4c)
- [Descrição e exemplos de validação de dados no Excel](https://support.microsoft.com/help/211485/description-and-examples-of-data-validation-in-excel)

## <a name="programmatic-control-of-data-validation"></a>Controle programático de validação de dados

A propriedade `Range.dataValidation`, que usa um objeto [DataValidation](https://docs.microsoft.com/javascript/api/excel/excel.datavalidation) , é o ponto de entrada para o controle programático da validação de dados no Excel. Há cinco propriedades para o objeto `DataValidation`:

- `rule` – Define o que constitui dados válidos para o intervalo. Confira [DataValidationRule](https://docs.microsoft.com/javascript/api/excel/excel.datavalidationrule).
- `errorAlert` – Especifica se um erro será exibido caso o usuário insira dados inválidos, e define o texto, o título e o estilo do alerta, por exemplo: **Informativo**, **Aviso**e **Parar**. Confira [DataValidationErrorAlert](https://docs.microsoft.com/javascript/api/excel/excel.datavalidationerroralert).
- `prompt` – Especifica se uma solicitação será exibida quando o usuário passar o mouse sobre o intervalo e define a mensagem da solicitação. Confira [DataValidationPrompt](https://docs.microsoft.com/javascript/api/excel/excel.datavalidationprompt).
- `ignoreBlanks` – Especifica se a regra de validação de dados se aplica a células em branco no intervalo. Padrão para `true`.
- `type` – Uma identificação somente leitura do tipo de validação, como WholeNumber, Date, TextLength etc. Ela é definida indiretamente ao se definir a propriedade `rule`.

> [!NOTE]
> A validação de dados adicionada de forma programática se comporta exatamente como manualmente adicionada a validação de dados. Em particular, observe que a validação de dados é acionada apenas se o usuário insere um valor em uma célula ou copia e cola uma célula de qualquer outro lugar na pasta de trabalho e escolhe  a opção de colagem **Valores**. Se o usuário copiar uma célula e fizer uma colagem sem formatação em um intervalo com validação de dados, a validação não será acionada.

## <a name="creating-validation-rules"></a>Criando regras de validação

Para adicionar a validação de dados a um intervalo, seu código deve definir a propriedade `rule` do objeto `DataValidation` em `Range.dataValidation`. Isso leva a um objeto [DataValidationRule](https://docs.microsoft.com/javascript/api/excel/excel.datavalidationrule) que tem sete propriedades opcionais. *Não mais de uma dessas propriedades pode estar presente em qualquer `DataValidationRule` objeto.* A propriedade que você incluir determina o tipo de validação.

### <a name="basic-and-datetime-validation-rule-types"></a>Tipos de regra de validação Basic e DateTime

As três primeiras propriedades `DataValidationRule` (isto é, tipos de regra de validação) usam um objeto [BasicDataValidation](https://docs.microsoft.com/javascript/api/excel/excel.basicdatavalidation) como seu valor.

- `wholeNumber` – Requer um número inteiro, além de qualquer outra validação especificada pelo objeto `BasicDataValidation`.
- `decimal` – Requer um número decimal, além de qualquer outra validação especificada pelo objeto `BasicDataValidation`.
- `textLength` – Aplica os detalhes de validação no objeto `BasicDataValidation` ao *comprimento* do valor da célula.

Este é um exemplo de como criar uma regra de validação. Observe o seguinte sobre este código:

- O `operator` é o operador binário "GreaterThan". Sempre que você usar um operador binário, o valor que o usuário tentar inserir na célula é o operando esquerdo e o valor especificado em `formula1` é o operando direito. Portanto, esta regra diz que apenas os números inteiros maiores que 0 são válidos. 
- O `formula1` é um número codificado. Se não souber o que valor deve ser no momento da codificação, você também pode usar uma fórmula do Excel (como uma sequência de caracteres) para o valor. Por exemplo, "= A3" e "= SUM(A4,B5)" também poderiam ser valores de `formula1`.

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

Consulte [BasicDataValidation](https://docs.microsoft.com/javascript/api/excel/excel.basicdatavalidation) para obter uma lista dos outros operadores binários. 

Também há dois operadores ternários: "Between" e "NotBetween". Para usá-los, é preciso especificar a propriedade opcional `formula2`. Os valores `formula1` e `formula2` são os operandos delimitadores. O valor que o usuário tenta inserir na célula é o terceiro operando (avaliado). Este é um exemplo de utilização do operador "Between":

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

As próximas duas propriedades da regra usam o objeto [DateTimeDataValidation](https://docs.microsoft.com/javascript/api/excel/excel.datetimedatavalidation) como seu valor.

- `date`
- `time`

O objeto `DateTimeDataValidation` é estruturado da mesma forma que o `BasicDataValidation`: ele tem as propriedades `formula1`, `formula2` e `operator` e é usado da mesma maneira. A diferença é que você não pode usar um número nas propriedades da fórmulas, mas pode inserir uma sequência de caracteres [ISO 8606 datetime](https://www.iso.org/iso-8601-date-and-time-format.html) (ou uma fórmula do Excel). Este é um exemplo que define os valores válidos como datas na primeira semana de abril de 2018. 

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

### <a name="list-validation-rule-type"></a>Tipo de regra de validação de lista

Use a propriedade `list` no objeto `DataValidationRule` para especificar que os únicos valores válidos são aqueles de uma lista finita. Este é um exemplo. Observe o seguinte sobre este código:

- Ele pressupõe que há uma planilha chamada “Nomes” e que os valores no intervalo “A1:A3” são nomes.
- A propriedade `source` especifica a lista de valores válidos. O intervalo com os nomes foi atribuído a ela. Você também pode atribuir uma lista delimitada por vírgula; por exemplo: "Sue, Ricky, Liz". 
- A propriedade `inCellDropDown` especifica se um controle da lista suspensa será exibido na célula quando o usuário o selecionar. Se for definido como `true`, a lista suspensa será exibida com a lista de valores de `source`.

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

### <a name="custom-validation-rule-type"></a>Tipo de regra de validação personalizada

Use a propriedade `custom` no objeto `DataValidationRule` para especificar uma fórmula de validação personalizada. Este é um exemplo. Observe o seguinte sobre este código:

- Ele pressupõe que há uma tabela de duas colunas com as colunas **Nome do Atleta** e **Comentários** nas colunas A e B da planilha.
- Para reduzir o detalhamento na coluna **Comentários**, ele faz com que os dados que incluem o nome do atleta se tornem inválidos.
- `SEARCH(A2,B2)` retorna a posição inicial, na sequência de caracteres em B2, da sequência de caracteres em A2. Se A2 não estiver contido em B2, ele não retorna um número. `ISNUMBER()` retorna um booleano. Portanto, a propriedade `formula` diz que os dados válidos para a coluna **Comentários** são dados que não incluem a sequência de caracteres na coluna **Nome do atleta** .

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

Você pode criar um alerta de erro personalizado que aparece quando um usuário tentar inserir dados inválidos em uma célula. Este é um exemplo simples. Observe o seguinte sobre este código:

- A propriedade `style` determina se o usuário obtém um alerta informativo, um aviso ou um alerta de "parar". Somente `Stop` pode realmente impedir que o usuário adicione dados inválidos. A janela pop-up para `Warning` e `Information` tem opções que permitem que o usuário insira dados inválidos mesmo assim.
- A propriedade `showAlert` assume `true` como padrão. Isso significa que o host do Excel irá abrir um alerta genérico (do tipo `Stop`), a menos que você crie um alerta personalizado que define `showAlert` para `false` ou define uma mensagem personalizada, título e estilo. Este código define uma mensagem personalizada e um título.


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

Para obter mais informações, confira [DataValidationErrorAlert](https://docs.microsoft.com/javascript/api/excel/excel.datavalidationerroralert).

## <a name="create-validation-prompts"></a>Criar solicitações de validação

É possível criar uma solicitação de instrução que aparece quando um usuário seleciona ou passa o mouse sobre uma célula na qual a validação de dados foi aplicada. Este é um exemplo:

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

Para obter mais informações, confira [DataValidationPrompt](https://docs.microsoft.com/javascript/api/excel/excel.datavalidationprompt).

## <a name="remove-data-validation-from-a-range"></a>Remover a validação de dados de um intervalo

Para remover a validação de dados de um intervalo, chame o método [Range.dataValidation.clear()](https://docs.microsoft.com/javascript/api/excel/excel.datavalidation#clear--).

```js
myrange.dataValidation.clear()
```

Não é necessário que o intervalo que você desmarcar seja exatamente o mesmo intervalo no qual você adicionou a validação de dados. Se não for, somente as células sobrepostas, se houver, de dois intervalos serão desmarcadas. 

> [!NOTE]
> A desmarcação da validação de dados de um intervalo também será aplicada em qualquer validação de dados que um usuário tiver adicionado manualmente ao intervalo.

## <a name="see-also"></a>Confira também

- [Conceitos fundamentais de programação com a API JavaScript do Excel](excel-add-ins-core-concepts.md)
- [Objeto DataValidation (API JavaScript para Excel)](https://docs.microsoft.com/javascript/api/excel/excel.datavalidation)
- [Objeto Range (API JavaScript para Excel)](https://docs.microsoft.com/javascript/api/excel/excel.range)



 
