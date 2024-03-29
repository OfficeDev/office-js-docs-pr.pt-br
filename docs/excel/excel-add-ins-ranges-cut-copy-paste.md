---
title: Intervalos de corte, cópia e colar usando a EXCEL JavaScript
description: Saiba como cortar, copiar e colar intervalos usando Excel API JavaScript.
ms.date: 02/16/2022
ms.prod: excel
ms.localizationpriority: medium
ms.openlocfilehash: 3d55e4d868a15c35ab9c68c799865560547e8188
ms.sourcegitcommit: 968d637defe816449a797aefd930872229214898
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 03/23/2022
ms.locfileid: "63745102"
---
# <a name="cut-copy-and-paste-ranges-using-the-excel-javascript-api"></a>Intervalos de corte, cópia e colar usando a EXCEL JavaScript

Este artigo fornece exemplos de código que cortam, copiam e colaram intervalos usando Excel API JavaScript. Para ver a lista completa de propriedades e métodos compatíveis `Range` com o objeto, [consulte Excel. Classe Range](/javascript/api/excel/excel.range).

[!include[Excel cells and ranges note](../includes/note-excel-cells-and-ranges.md)]

## <a name="copy-and-paste"></a>Copy and paste

O [método Range.copyFrom](/javascript/api/excel/excel.range#excel-excel-range-copyfrom-member(1)) replica as ações **Copiar** e **Colar** da interface Excel interface do usuário. O destino é o `Range` objeto chamado `copyFrom` . A fonte a ser copiada é passada como um intervalo ou um endereço de cadeia de caracteres que representa um intervalo.

O exemplo a seguir copia dados de **A1:E1** para o intervalo que começa em **G1** (que acaba sendo colado em **G1:K1**).

```js
await Excel.run(async (context) => {
    let sheet = context.workbook.worksheets.getItem("Sample");
    // Copy everything from "A1:E1" into "G1" and the cells afterwards ("G1:K1").
    sheet.getRange("G1").copyFrom("A1:E1");
    await context.sync();
});
```

`Range.copyFrom` tem três parâmetros opcionais.

```TypeScript
copyFrom(sourceRange: Range | RangeAreas | string, copyType?: Excel.RangeCopyType, skipBlanks?: boolean, transpose?: boolean): void;
```

`copyType` especifica quais dados são copiados da origem para o destino.

- `Excel.RangeCopyType.formulas` transfere as fórmulas nas células de origem e preserva o posicionamento relativo dos intervalos dessas fórmulas. As entradas que não sejam uma fórmula são copiadas no seu estado original.
- `Excel.RangeCopyType.values` copia os valores dos dados e, no caso de fórmulas, o resultado da fórmula.
- `Excel.RangeCopyType.formats` copia a formatação do intervalo incluindo a fonte, cor e outras configurações de formato, mas nenhum valor.
- `Excel.RangeCopyType.all` (a opção padrão) copia os dados e a formatação, preservando as fórmulas das células, se encontradas.

`skipBlanks` define se as células em branco são copiadas para o destino. Quando for verdadeiro, `copyFrom` ignora células em branco no intervalo de origem.
As células ignoradas não substituem os dados existentes de suas células correspondentes no intervalo de destino. O padrão é false.

`transpose` determina se os dados são transpostos, ou seja, suas linhas e colunas são alternadas para o local de origem.
Um intervalo transposto invertido na diagonal principal, portanto as linhas **1**, **2** e **3** se tornarão as colunas **A**, **B** e **C**.

O exemplo de código e as imagens a seguir demonstram esse comportamento em um cenário simples.

```js
await Excel.run(async (context) => {
    let sheet = context.workbook.worksheets.getItem("Sample");
    // Copy a range, omitting the blank cells so existing data is not overwritten in those cells.
    sheet.getRange("D1").copyFrom("A1:C1",
        Excel.RangeCopyType.all,
        true, // skipBlanks
        false); // transpose
    // Copy a range, including the blank cells which will overwrite existing data in the target cells.
    sheet.getRange("D2").copyFrom("A2:C2",
        Excel.RangeCopyType.all,
        false, // skipBlanks
        false); // transpose
    await context.sync();
});
```

### <a name="data-before-range-is-copied-and-pasted"></a>Dados antes que o intervalo seja copiado e passado

![Dados em Excel antes que o método de cópia do intervalo tenha sido executado.](../images/excel-range-copyfrom-skipblanks-before.png)

### <a name="data-after-range-is-copied-and-pasted"></a>Dados após o intervalo são copiados e copiados

![Dados na Excel após a executar o método de cópia do intervalo.](../images/excel-range-copyfrom-skipblanks-after.png)

## <a name="cut-and-paste-move-cells"></a>Cortar e colar células (mover)

O [método Range.moveTo](/javascript/api/excel/excel.range#excel-excel-range-moveto-member(1)) move células para um novo local na workbook. Esse comportamento de movimento de célula funciona da mesma forma que quando as [](https://support.microsoft.com/office/803d65eb-6a3e-4534-8c6f-ff12d1c4139e) células são movidas arrastando a borda do intervalo ou ao tomar as ações **Cortar** **e Colar**. Tanto a formatação quanto os valores do intervalo são movidos para o local especificado como o `destinationRange` parâmetro.

O exemplo de código a seguir move um intervalo com o `Range.moveTo` método. Observe que, se o intervalo de destino for menor que a fonte, ele será expandido para abranger o conteúdo de origem.

```js
await Excel.run(async (context) => {
    let sheet = context.workbook.worksheets.getActiveWorksheet();
    sheet.getRange("F1").values = [["Moved Range"]];

    // Move the cells "A1:E1" to "G1" (which fills the range "G1:K1").
    sheet.getRange("A1:E1").moveTo("G1");
    await context.sync();
});
```

## <a name="see-also"></a>Confira também

- [Modelo de objeto JavaScript do Excel em Suplementos do Office](excel-add-ins-core-concepts.md)
- [Trabalhar com células usando a EXCEL JavaScript](excel-add-ins-cells.md)
- [Remover duplicatas usando a EXCEL JavaScript](excel-add-ins-ranges-remove-duplicates.md)
- [Trabalhar simultaneamente com vários intervalos em suplementos do Excel](excel-add-ins-multiple-ranges.md)
