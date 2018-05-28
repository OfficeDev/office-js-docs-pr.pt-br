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
# <a name="add-data-validation-to-excel-ranges-preview"></a>Adicionar valida??o de dados a intervalos do Excel (vers?o pr?via)

> [!NOTE]
> Enquanto as APIs de valida??o de dados est?o em vers?o pr?via, voc? deve carregar a vers?o beta da biblioteca JavaScript do Office para us?-las. A URL ? https://appsforoffice.microsoft.com/lib/beta/hosted/office.js. Se voc? estiver usando o TypeScript ou se seu editor de c?digo usa um arquivo de defini??o do tipo TypeScript para IntelliSense, use https://appsforoffice.microsoft.com/lib/beta/hosted/office.d.ts.

A Biblioteca JavaScript do Excel fornece APIs para permitir que seu suplemento adicione valida??o de dados autom?tica a tabelas, colunas, linhas e outros intervalos em uma pasta de trabalho. Para entender os conceitos e a terminologia de valida??o de dados, consulte os artigos a seguir sobre como os usu?rios adicionam valida??o de dados por meio da interface do usu?rio do Excel:

- [Aplicar valida??o de dados a c?lulas](https://support.office.com/en-us/article/Apply-data-validation-to-cells-29FECBCC-D1B9-42C1-9D76-EFF3CE5F7249)
- [Mais sobre valida??o de dados](https://microsoft.sharepoint.com/:p:/r/teams/oext/_layouts/15/Doc.aspx?sourcedoc=%7B51143964-d52c-429d-bfac-c7495473d536%7D&action=edit)
- [Descri??o e exemplos de valida??o de dados no Excel](https://support.microsoft.com/en-us/help/211485/description-and-examples-of-data-validation-in-excel)

## <a name="programmatic-control-of-data-validation"></a>Controle program?tico de valida??o de dados

A propriedade`Range.dataValidation`, a qual usa um objeto[DataValidation](https://dev.office.com/reference/add-ins/excel/datavalidation), ? o ponto de entrada para o controle program?tico de valida??o de dados no Excel. Existem cinco propriedades para o objeto `DataValidation`:

- `rule` ? Define o que constitui dados v?lidos para o intervalo. Consulte [DataValidationRule](https://dev.office.com/reference/add-ins/excel/datavalidationrule).
- `errorAlert` ? Especifica se um erro ser? exibido caso o usu?rio insira dados inv?lidos e define o texto, o t?tulo e o estilo do alerta, por exemplo: **Informativo**, **Aten??o**e **Pare**. Consulte [DataValidationErrorAlert](https://dev.office.com/reference/add-ins/excel/datavalidationerroralert).
- `prompt` ? Especifica se uma solicita??o ser? exibida quando o usu?rio passar o mouse sobre o intervalo e define a mensagem da solicita??o. Consulte [DataValidationPrompt](https://dev.office.com/reference/add-ins/excel/datavalidationprompt).
- `ignoreBlanks` ? Especifica se a regra de valida??o de dados se aplica a c?lulas em branco no intervalo. Padr?es para `true`.
- `type` ? Uma identifica??o somente leitura do tipo de valida??o, como WholeNumber, Date, TextLength, etc. Ela ? definida indiretamente ao se definir a propriedade `rule`.

> [!NOTE]
> A valida??o de dados adicionada programaticamente se comporta exatamente como a valida??o de dados adicionada manualmente. Em particular, observe que a valida??o de dados s? ? acionada se o usu?rio inserir um valor diretamente em uma c?lula ou copiar e colar uma c?lula de outro local da pasta de trabalho e escolher a op??o de colar**Valores**. Se o usu?rio copiar uma c?lula e executar a a??o de colar sem formata??o em um intervalo com valida??o de dados, a valida??o n?o ser? acionada.

### <a name="creating-validation-rules"></a>Criando regras de valida??o

Para adicionar valida??o de dados a um intervalo, seu c?digo deve definir propriedade `rule` do objeto `DataValidation` em `Range.dataValidation`. Usa-se um objeto [DataValidationRule](https://dev.office.com/reference/add-ins/excel/datavalidationrule) que tem sete propriedades opcionais. *N?o pode haver mais do que uma dessas propriedades presente em qualquer objeto `DataValidationRule`.* A propriedade inclu?da por voc? determina o tipo de valida??o.

#### <a name="basic-and-datetime-validation-rule-types"></a>Tipos de regra de valida??o B?sico e DateTime

As tr?s primeiras propriedades `DataValidationRule` (isto ?, tipos de regra de valida??o) usam um objeto [BasicDataValidation](https://dev.office.com/reference/add-ins/excel/basicdatavalidation) como seu valor.

- `wholeNumber` ? Requer um n?mero inteiro, al?m de qualquer outra valida??o especificada pelo objeto `BasicDataValidation`.
- `decimal` ? Requer um n?mero decimal, al?m de qualquer outra valida??o especificada pelo objeto `BasicDataValidation`.
- `textLength` ? Aplica os detalhes de valida??o no objeto `BasicDataValidation` ao *comprimento* do valor da c?lula.

Este ? um exemplo de cria??o de uma regra de valida??o. Sobre este c?digo, observe o seguinte:

- O `operator`  ? o operador bin?rio ?GreaterThan?. Sempre for usado um operador bin?rio, o valor que o usu?rio tentar inserir na c?lula ? o operando esquerdo, e o valor especificado em `formula1` ? o operando direito. Portanto, essa regra diz que apenas n?meros inteiros maiores que 0 s?o v?lidos. 
- O `formula1` ? um n?mero embutido em c?digo. No momento da codifica??o, caso n?o saiba qual deve ser o valor, voc? tamb?m poder? usar uma f?rmula do Excel (como uma sequ?ncia de caracteres) para o valor. Por exemplo, ?= A3? e ?= SUM(A4, B5)? tamb?m podem ser valores de `formula1`.

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

Consulte [BasicDataValidation](https://dev.office.com/reference/add-ins/excel/basicdatavalidation) para obter uma lista dos outros operadores bin?rios. 

Existem tamb?m dois operadores tern?rios: ?Between? e ?NotBetween?. Para us?-los, voc? deve especificar a propriedade opcional `formula2`. Os valores `formula1` e `formula2` s?o os operandos delimitadores. O valor que o usu?rio tentar inserir na c?lula ? o terceiro operando (avaliado). A seguir, h? um exemplo de uso do operador ?Between?:

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

As pr?ximas duas propriedades da regra usam o objeto [DateTimeDataValidation](https://dev.office.com/reference/add-ins/excel/basicdatavalidation) como seu valor.

- `date`
- `time`

O objeto `DateTimeDataValidation` ? estruturado de forma semelhante ao `BasicDataValidation`: tem as propriedades `formula1`, `formula2`e `operator` e ? usado da mesma maneira. A diferen?a ? que voc? n?o pode usar um n?mero nas propriedades da f?rmula, mas pode inserir uma sequ?ncia de caracteres [ISO 8606 datetime](https://www.iso.org/iso-8601-date-and-time-format.html) (ou uma f?rmula do Excel). A seguir, h? um exemplo que define valores v?lidos como datas na primeira semana de abril de 2018. 

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

#### <a name="list-validation-rule-type"></a>Tipo de regra de valida??o de lista

Use a propriedade `list` no objeto `DataValidationRule` para especificar que os ?nicos valores v?lidos sejam aqueles de uma lista finita. H? um exemplo a seguir. Sobre este c?digo, observe o seguinte:

- Ele pressup?e que h? uma planilha chamada ?Nomes? e que os valores no intervalo ?A1: A3? s?o nomes.
- A propriedade `source` especifica a lista de valores v?lidos. O intervalo com os nomes foi atribu?do a ela. Tamb?m ? poss?vel atribuir uma lista delimitada por v?rgula, por exemplo: ?Sue, Ricky, Liz?. 
- A propriedade `inCellDropDown` especifica se um controle suspenso aparecer? na c?lula quando o usu?rio selecion?-lo. Se definida como `true`, a lista suspensa aparece com a lista de valores de `source`.

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

#### <a name="custom-validation-rule-type"></a>Tipo de regra de valida??o personalizada

Use a propriedade `custom` no objeto `DataValidationRule` para especificar uma f?rmula de valida??o personalizada. H? um exemplo a seguir. Sobre este c?digo, observe o seguinte:

- Ele pressup?e que h? uma tabela de duas colunas com colunas **Nome do Atleta** e **Coment?rios** nas colunas A e B da planilha.
- Para reduzir a verbosidade na coluna **Coment?rios**, ele faz com que os dados que incluem o nome do atleta se tornem inv?lidos.
- `SEARCH(A2,B2)` retorna a posi??o inicial, na sequ?ncia de caracteres B2, da sequ?ncia de caracteres em A2. Se A2 n?o estiver contida em B2, ele n?o retornar? um n?mero. `ISNUMBER()` retorna um booleano. Ent?o a propriedade `formula` diz que dados v?lidos da coluna **Coment?rios** s?o dados que n?o incluem a sequ?ncia de caracteres na coluna **Nome do Atleta**.

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

### <a name="create-validation-error-alerts"></a>Criar alertas de erro de valida??o

? poss?vel criar um alerta de erro personalizado que aparecer? quando um usu?rio tentar inserir dados inv?lidos em uma c?lula. H? um exemplo simples a seguir. Sobre este c?digo, observe o seguinte:

- A propriedade `style` determina se o usu?rio recebe um alerta informativo, um aviso ou um alerta do tipo ?pare?. Somente `Stop` impede de verdade que o usu?rio adicione dados inv?lidos. O pop-up para `Warning` e `Information` tem op??es que permitem que o usu?rio insira os dados inv?lidos.
- A propriedade `showAlert` se torna padr?o para `true`. Isso significa que o host do Excel exibir? um alerta pop-up gen?rico (do tipo `Stop`) a menos que seja criado um alerta personalizado que defina `showAlert` para `false` ou defina uma mensagem, um t?tulo e um estilo personalizados. Esse c?digo define uma mensagem personalizada e um t?tulo.


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

Para obter mais informa??es, consulte [DataValidationErrorAlert](https://dev.office.com/reference/add-ins/excel/datavalidationerroralert).

### <a name="create-validation-prompts"></a>Criar solicita??es de valida??o

? poss?vel criar uma solicita??o de instru??o que aparece quando um usu?rio seleciona ou passa o mouse sobre uma c?lula na qual a valida??o de dados foi aplicada. H? um exemplo a seguir:

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

Para obter mais informa??es, consulte [DataValidationPrompt](https://dev.office.com/reference/add-ins/excel/datavalidationprompt).

### <a name="remove-data-validation-from-a-range"></a>Remover a valida??o de dados de um intervalo

Para remover a valida??o de dados de um intervalo, chame o m?todo [Range.dataValidation.clear()](https://dev.office.com/reference/add-ins/excel/datavalidation#clear).

```js
myrange.dataValidation.clear()
```

N?o ? necess?rio que o intervalo limpo seja exatamente o mesmo de um intervalo no qual a valida??o de dados foi adicionada. Se n?o for, apenas as c?lulas sobrepostas, se houver, dos dois intervalos ser?o limpas. 

> [!NOTE]
> A limpeza da valida??o de dados de um intervalo tamb?m limpar? qualquer valida??o de dados que um usu?rio tenha adicionado manualmente ao intervalo.

## <a name="see-also"></a>Confira tamb?m

- [Principais conceitos da API JavaScript do Excel](excel-add-ins-core-concepts.md)
- [Objeto DataValidation (API JavaScript para Excel)](https://dev.office.com/reference/add-ins/excel/datavalidation)
- [Objeto Range (API JavaScript para Excel)](https://dev.office.com/reference/add-ins/excel/range)



 
