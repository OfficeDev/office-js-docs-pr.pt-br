---
title: Trabalhar com intervalos usando a API JavaScript do Excel (avançado)
description: ''
ms.date: 12/14/2018
ms.openlocfilehash: 42b1127580c46120d337553fdb86a19a78b37567
ms.sourcegitcommit: 09f124fac7b2e711e1a8be562a99624627c0699e
ms.translationtype: HT
ms.contentlocale: pt-BR
ms.lasthandoff: 12/15/2018
ms.locfileid: "27283790"
---
# <a name="work-with-ranges-using-the-excel-javascript-api-advanced"></a>Trabalhar com intervalos usando a API JavaScript do Excel (avançado)

Este artigo baseia-se em informações em [Trabalhar com intervalos usando a API JavaScript do Excel (fundamental)](excel-add-ins-ranges.md) fornecendo exemplos de código que mostram como executar tarefas mais avançadas com intervalos usando a API JavaScript do Excel. Para obter a lista completa de propriedades e métodos que o objeto **Range** suporta, confira [Objeto Range (API JavaScript para Excel)](https://docs.microsoft.com/javascript/api/excel/excel.range).

## <a name="work-with-dates-using-the-moment-msdate-plug-in"></a>Trabalhar com datas usando o plug-in Moment-MSDate

A [biblioteca Moment do JavaScript](https://momentjs.com/) fornece uma maneira conveniente de usar datas e carimbos de data e hora. O [plug-in Moment-MSDate](https://www.npmjs.com/package/moment-msdate) converte o formato de momentos em um formato mais apropriado para o Excel. Este é o mesmo formato que a [função NOW](https://support.office.com/article/now-function-3337fd29-145a-4347-b2e6-20c904739c46) retorna.

O código a seguir mostra como definir o intervalo em ** B4 ** para o carimbo de data/hora de um momento:

```js
Excel.run(function (context) {
    var sheet = context.workbook.worksheets.getItem("Sample");

    var now = Date.now();
    var nowMoment = moment(now);
    var nowMS = nowMoment.toOADate();

    var dateRange = sheet.getRange("B4");
    dateRange.values = [[nowMS]];

    dateRange.numberFormat = [["[$-409]m/d/yy h:mm AM/PM;@"]];

    return context.sync();
}).catch(errorHandlerFunction);
```

É uma técnica semelhante para retirar a data da célula e convertê-la em um momento ou outro formato, conforme demonstrado no código a seguir:

```js
Excel.run(function (context) {
    var sheet = context.workbook.worksheets.getItem("Sample");

    var dateRange = sheet.getRange("B4");
    dateRange.load("values");

    return context.sync().then(function () {
        var nowMS = dateRange.values[0][0];

        // log the date as a moment
        var nowMoment = moment.fromOADate(nowMS);
        console.log(`get (moment): ${JSON.stringify(nowMoment)}`);

        // log the date as a UNIX-style timestamp 
        var now = nowMoment.unix();
        console.log(`get (timestamp): ${now}`);
    });
}).catch(errorHandlerFunction);
```

Seu suplemento terá que formatar os intervalos para exibir as datas em um formato mais legível. O exemplo de `"[$-409]m/d/yy h:mm AM/PM;@"` exibe a hora como "3/12/18 15:57". Para obter mais informações sobre formatos de números de data e hora, confira as "Diretrizes para formatos de data e hora" no artigo [Diretrizes de revisão para personalizar um formato de número](https://support.office.com/article/review-guidelines-for-customizing-a-number-format-c0a1d1fa-d3f4-4018-96b7-9c9354dd99f5).

## <a name="copy-and-paste"></a>Copiar e colar

> [!NOTE]
> A função`Range.copyFrom` no momento só está disponível na versão prévia pública (beta). Para usar esse recurso, você deve usar a biblioteca beta do CDN do Office.js: https://appsforoffice.microsoft.com/lib/beta/hosted/office.js.
> Se você estiver usando o TypeScript ou se seu editor de código usar arquivos de definição de tipo do TypeScript do IntelliSense, use https://appsforoffice.microsoft.com/lib/beta/hosted/office.d.ts.

A função de `copyFrom` do intervalo replica o comportamento de copiar e colar da IU do Excel. O objeto de intervalo para o qual a função`copyFrom` é chamada é o destino.
A fonte a ser copiada é passada como um intervalo ou um endereço de cadeia de caracteres que representa um intervalo. O exemplo a seguir copia dados de **A1:E1** para o intervalo que começa em **G1** (que acaba sendo colado em **G1:K1**).

```js
Excel.run(function (context) {
    var sheet = context.workbook.worksheets.getItem("Sample");
    // copy a range starting at a single cell destination
    sheet.getRange("G1").copyFrom("A1:E1");
    return context.sync();
}).catch(errorHandlerFunction);
```

`Range.copyFrom` tem três parâmetros opcionais.

```TypeScript
copyFrom(sourceRange: Range | string, copyType?: "All" | "Formulas" | "Values" | "Formats", skipBlanks?: boolean, transpose?: boolean): void;
```

`copyType` especifica quais dados são copiados da origem para o destino.

- `"Formulas"` transfere as fórmulas nas células de origem e preserva o posicionamento relativo dos intervalos dessas fórmulas. As entradas que não sejam uma fórmula são copiadas no seu estado original.
- `"Values"` copia os valores dos dados e, no caso de fórmulas, o resultado da fórmula.
- `"Formats"` copia a formatação do intervalo incluindo a fonte, cor e outras configurações de formato, mas nenhum valor.
- `"All"` (a opção padrão) copia ambos os dados e formatação, preservando as fórmulas das células, caso elas sejam encontradas.

`skipBlanks` define se as células em branco são copiadas para o destino. Quando for verdadeiro, `copyFrom` ignora células em branco no intervalo de origem.
As células ignoradas não substituem os dados existentes de suas células correspondentes no intervalo de destino. O padrão é false.

`transpose` determina se os dados são transpostos, ou seja, suas linhas e colunas são alternadas para o local de origem.
Um intervalo transposto invertido na diagonal principal, portanto as linhas **1**, **2** e **3** se tornarão as colunas **A**, **B** e **C**.

O exemplo de código e as imagens a seguir demonstram esse comportamento em um cenário simples.

```js
Excel.run(function (context) {
    var sheet = context.workbook.worksheets.getItem("Sample");
    // copy a range, omitting the blank cells so existing data is not overwritten in those cells
    sheet.getRange("D1").copyFrom("A1:C1",
        Excel.RangeCopyType.all,
        true, // skipBlanks
        false); // transpose
    // copy a range, including the blank cells which will overwrite existing data in the target cells
    sheet.getRange("D2").copyFrom("A2:C2",
        Excel.RangeCopyType.all,
        false, // skipBlanks
        false); // transpose
    return context.sync();
}).catch(errorHandlerFunction);
```

*Antes da função precedente ter sido executada.*

![Os dados no Excel antes do método de copiar do intervalo foram executados](../images/excel-range-copyfrom-skipblanks-before.png)

*Após a função precedente ter sido executada.*

![Os dados no Excel após o método de copiar do intervalo foram executados](../images/excel-range-copyfrom-skipblanks-after.png)

## <a name="remove-duplicates"></a>Remover duplicatas

> [!NOTE]
> A função `removeDuplicates` no momento só está disponível na versão prévia pública (beta). Para usar esse recurso, você deve usar a biblioteca beta do CDN do Office.js: https://appsforoffice.microsoft.com/lib/beta/hosted/office.js.
> Se você estiver usando o TypeScript ou se seu editor de código usar arquivos de definição de tipo do TypeScript do IntelliSense, use https://appsforoffice.microsoft.com/lib/beta/hosted/office.d.ts.

A função do objeto intervalo `removeDuplicates` remove linhas com entradas duplicadas em determinadas colunas. A função passa por cada linha no intervalo do índice de menor valor até o índice de maior valor no intervalo (de cima para baixo). Uma linha é excluída se um valor em sua coluna ou colunas especificadas aparecer mais cedo no intervalo. Linhas no intervalo abaixo da linha excluída são deslocadas para cima. `removeDuplicates` não afeta a posição de células fora do intervalo.

`removeDuplicates` leva um `number[]` representando os índices da coluna que são verificados para duplicatas. Essa matriz é baseada em zero e relativa ao intervalo, não à planilha. A função também aceita um parâmetro booliano que especifica se a primeira linha é um cabeçalho. Quando **verdadeiro**, a primeira linha será ignorada ao considerar duplicatas. A função `removeDuplicates` retorna um objeto `RemoveDuplicatesResult` que especifica o número de linhas removidas e o número de linhas exclusivas restantes.

Ao usar um intervalo na função`removeDuplicates`, lembre-se do seguinte:

- `removeDuplicates` considera valores de célula, não resultados de função. Se as duas funções diferentes forem avaliadas como o mesmo resultado, os valores de célula não são considerados duplicatas.
- Células vazias não serão ignoradas por `removeDuplicates`. O valor de uma célula vazia é tratado como qualquer outro valor. Isso significa que as linhas vazias contidas no intervalo serão incluídas em `RemoveDuplicatesResult`.

O exemplo a seguir mostra a remoção de entradas com valores duplicados na primeira coluna.

```js
Excel.run(async (context) => {
    var sheet = context.workbook.worksheets.getItem("Sample");
    var range = sheet.getRange("B2:D11");

    var deleteResult = range.removeDuplicates([0],true);
    deleteResult.load();

    return context.sync().then(function () {
        console.log(deleteResult.removed + " entries with duplicate names removed.");
        console.log(deleteResult.uniqueRemaining + " entries with unique names remain in the range.");
    });
}).catch(errorHandlerFunction);
```

*Antes da função precedente ter sido executada.*

![Dados no Excel antes da execução do método de remoção de duplicatas do intervalo](../images/excel-ranges-remove-duplicates-before.png)

*Após a função precedente ter sido executada.*

![Dados no Excel depois da execução do método de remoção de duplicatas do intervalo](../images/excel-ranges-remove-duplicates-after.png)

## <a name="see-also"></a>Confira também

- [Trabalhar com intervalos usando a API JavaScript do Excel](excel-add-ins-ranges.md)
- [Conceitos fundamentais de programação com a API JavaScript do Excel](excel-add-ins-core-concepts.md)