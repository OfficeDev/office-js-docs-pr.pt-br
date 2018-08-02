---
title: Adicionar validação de dados a intervalos do Excel
description: ''
ms.date: 04/13/2018
ms.openlocfilehash: 3d6a901e2f8296806cff470340b40f4d77e79e34
ms.sourcegitcommit: bc68b4cf811b45e8b8d1cbd7c8d2867359ab671b
ms.translationtype: HT
ms.contentlocale: pt-BR
ms.lasthandoff: 08/02/2018
ms.locfileid: "21703942"
---
# <a name="add-data-validation-to-excel-ranges-preview"></a>Adicionar validação de dados a intervalos do Excel (versão prévia)

> [!NOTE]
> Enquanto as APIs de validação de dados estão em versão prévia, você deve carregar a versão beta da biblioteca JavaScript do Office para usá-las. A URL é https://appsforoffice.microsoft.com/lib/beta/hosted/office.js. Se você estiver usando o TypeScript ou se seu editor de código usa um arquivo de definição do tipo TypeScript para IntelliSense, use https://appsforoffice.microsoft.com/lib/beta/hosted/office.d.ts.

> [!NOTE]
> Embora as APIs de validação de dados estejam em versão prévia, os links neste artigo para a referência da API não funcionarão. Enquanto isso, você pode usar a [referência da API do Excel de rascunho](https://github.com/OfficeDev/office-js-docs/tree/ExcelJs_OpenSpec/reference/excel).

A Biblioteca JavaScript do Excel fornece APIs para permitir que seu suplemento adicione validação de dados automática a tabelas, colunas, linhas e outros intervalos em uma pasta de trabalho. Para entender os conceitos e a terminologia de validação de dados, consulte os artigos a seguir sobre como os usuários adicionam validação de dados por meio da interface do usuário do Excel:

- [Aplicar validação de dados a células](https://support.office.com/en-us/article/Apply-data-validation-to-cells-29FECBCC-D1B9-42C1-9D76-EFF3CE5F7249)
- [Mais sobre validação de dados](https://support.office.com/en-us/article/More-on-data-validation-f38dee73-9900-4ca6-9301-8a5f6e1f0c4c)
- [Descrição e exemplos de validação de dados no Excel](https://support.microsoft.com/en-us/help/211485/description-and-examples-of-data-validation-in-excel)

## <a name="programmatic-control-of-data-validation"></a>Controle programático de validação de dados

A propriedade`Range.dataValidation`, a qual usa um objeto[DataValidation](https://docs.microsoft.com/javascript/api/excel/excel.datavalidation), é o ponto de entrada para o controle programático de validação de dados no Excel. Existem cinco propriedades para o objeto `DataValidation`:

- `rule` – Define o que constitui dados válidos para o intervalo. Consulte [DataValidationRule](https://docs.microsoft.com/javascript/api/excel/excel.datavalidationrule).
- `errorAlert` – Especifica se um erro será exibido caso o usuário insira dados inválidos e define o texto, o título e o estilo do alerta, por exemplo: **Informativo**, **Atenção**e **Pare**. Consulte [DataValidationErrorAlert](https://docs.microsoft.com/javascript/api/excel/excel.datavalidationerroralert).
- `prompt` – Especifica se uma solicitação será exibida quando o usuário passar o mouse sobre o intervalo e define a mensagem da solicitação. Consulte [DataValidationPrompt](https://docs.microsoft.com/javascript/api/excel/excel.datavalidationprompt).
- `ignoreBlanks` – Especifica se a regra de validação de dados se aplica a células em branco no intervalo. Padrões para `true`.
- `type` – Uma identificação somente leitura do tipo de validação, como WholeNumber, Date, TextLength, etc. Ela é definida indiretamente ao se definir a propriedade `rule`.

> [!NOTE]
> A validação de dados adicionada programaticamente se comporta exatamente como a validação de dados adicionada manualmente. Em particular, observe que a validação de dados só é acionada se o usuário inserir um valor diretamente em uma célula ou copiar e colar uma célula de outro local da pasta de trabalho e escolher a opção de colar**Valores**. Se o usuário copiar uma célula e executar a ação de colar sem formatação em um intervalo com validação de dados, a validação não será acionada.

### <a name="creating-validation-rules"></a>Criando regras de validação

Para adicionar validação de dados a um intervalo, seu código deve definir propriedade `rule` do objeto `DataValidation` em `Range.dataValidation`. Usa-se um objeto [DataValidationRule](https://dev.office.com/reference/add-ins/excel/datavalidationrule) que tem sete propriedades opcionais. *Não pode haver mais do que uma dessas propriedades presente em qualquer objeto `DataValidationRule`.* A propriedade incluída por você determina o tipo de validação.

#### <a name="basic-and-datetime-validation-rule-types"></a>Tipos de regra de validação Básico e DateTime

As três primeiras propriedades `DataValidationRule` (isto é, tipos de regra de validação) usam um objeto [BasicDataValidation](https://docs.microsoft.com/javascript/api/excel/excel.basicdatavalidation) como seu valor.

- `wholeNumber` – Requer um número inteiro, além de qualquer outra validação especificada pelo objeto `BasicDataValidation`.
- `decimal` – Requer um número decimal, além de qualquer outra validação especificada pelo objeto `BasicDataValidation`.
- `textLength` – Aplica os detalhes de validação no objeto `BasicDataValidation` ao *comprimento* do valor da célula.

Este é um exemplo de criação de uma regra de validação. Observe o seguinte sobre este código:

- O `operator`  é o operador binário “GreaterThan”. Sempre for usado um operador binário, o valor que o usuário tentar inserir na célula é o operando esquerdo, e o valor especificado em `formula1` é o operando direito. Portanto, essa regra diz que apenas números inteiros maiores que 0 são válidos. 
- O `formula1` é um número embutido em código. No momento da codificação, caso não saiba qual deve ser o valor, você também poderá usar uma fórmula do Excel (como uma sequência de caracteres) para o valor. Por exemplo, “= A3” e “= SUM(A4, B5)” também podem ser valores de `formula1`.

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

Existem também dois operadores ternários: “Between” e “NotBetween”. Para usá-los, você deve especificar a propriedade opcional `formula2`. Os valores `formula1` e `formula2` são os operandos delimitadores. O valor que o usuário tentar inserir na célula é o terceiro operando (avaliado). A seguir, há um exemplo de uso do operador “Between”:

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

O objeto `DateTimeDataValidation` é estruturado de forma semelhante ao `BasicDataValidation`: tem as propriedades `formula1`, `formula2`e `operator` e é usado da mesma maneira. A diferença é que você não pode usar um número nas propriedades da fórmula, mas pode inserir uma sequência de caracteres [ISO 8606 datetime](https://www.iso.org/iso-8601-date-and-time-format.html) (ou uma fórmula do Excel). A seguir, há um exemplo que define valores válidos como datas na primeira semana de abril de 2018. 

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

#### <a name="list-validation-rule-type"></a>Tipo de regra de validação de lista

Use a propriedade `list` no objeto `DataValidationRule` para especificar que os únicos valores válidos sejam aqueles de uma lista finita. Apresentamos um exemplo a seguir. Observe o seguinte sobre este código:

- Ele pressupõe que há uma planilha chamada “Nomes” e que os valores no intervalo “A1: A3” são nomes.
- A propriedade `source` especifica a lista de valores válidos. O intervalo com os nomes foi atribuído a ela. Também é possível atribuir uma lista delimitada por vírgula, por exemplo: “Sue, Ricky, Liz”. 
- A propriedade `inCellDropDown` especifica se um controle suspenso aparecerá na célula quando o usuário selecioná-lo. Se definida como `true`, a lista suspensa aparece com a lista de valores de `source`.

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

#### <a name="custom-validation-rule-type"></a>Tipo de regra de validação personalizada

Use a propriedade `custom` no objeto `DataValidationRule` para especificar uma fórmula de validação personalizada. Apresentamos um exemplo a seguir. Observe o seguinte sobre este código:

- Ele pressupõe que há uma tabela de duas colunas com colunas **Nome do Atleta** e **Comentários** nas colunas A e B da planilha.
- Para reduzir a verbosidade na coluna **Comentários**, ele faz com que os dados que incluem o nome do atleta se tornem inválidos.
- `SEARCH(A2,B2)` retorna a posição inicial, na sequência de caracteres B2, da sequência de caracteres em A2. Se A2 não estiver contida em B2, ele não retornará um número. `ISNUMBER()` retorna um booleano. Então a propriedade `formula` diz que dados válidos da coluna **Comentários** são dados que não incluem a sequência de caracteres na coluna **Nome do Atleta**.

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

### <a name="create-validation-error-alerts"></a>Criar alertas de erro de validação

É possível criar um alerta de erro personalizado que aparecerá quando um usuário tentar inserir dados inválidos em uma célula. Há um exemplo simples a seguir. Observe o seguinte sobre este código:

- A propriedade `style` determina se o usuário recebe um alerta informativo, um aviso ou um alerta do tipo “pare”. Somente `Stop` impede de verdade que o usuário adicione dados inválidos. O pop-up para `Warning` e `Information` tem opções que permitem que o usuário insira os dados inválidos.
- A propriedade `showAlert` se torna padrão para `true`. Isso significa que o host do Excel exibirá um alerta pop-up genérico (do tipo `Stop`) a menos que seja criado um alerta personalizado que defina `showAlert` para `false` ou defina uma mensagem, um título e um estilo personalizados. Esse código define uma mensagem personalizada e um título.


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

### <a name="create-validation-prompts"></a>Criar solicitações de validação

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

### <a name="remove-data-validation-from-a-range"></a>Remover a validação de dados de um intervalo

Para remover a validação de dados de um intervalo, chame o método [Range.dataValidation.clear()](https://docs.microsoft.com/javascript/api/excel/excel.datavalidation#clear).

```js
myrange.dataValidation.clear()
```

Não é necessário que o intervalo limpo seja exatamente o mesmo de um intervalo no qual a validação de dados foi adicionada. Se não for, apenas as células sobrepostas, se houver, dos dois intervalos serão limpas. 

> [!NOTE]
> A limpeza da validação de dados de um intervalo também limpará qualquer validação de dados que um usuário tenha adicionado manualmente ao intervalo.

## <a name="see-also"></a>Veja também

- [Principais conceitos da API JavaScript do Excel](excel-add-ins-core-concepts.md)
- [Objeto DataValidation (API JavaScript para Excel)](https://docs.microsoft.com/javascript/api/excel/excel.datavalidation)
- [Objeto Range (API JavaScript para Excel)](https://docs.microsoft.com/javascript/api/excel/excel.range)



 
