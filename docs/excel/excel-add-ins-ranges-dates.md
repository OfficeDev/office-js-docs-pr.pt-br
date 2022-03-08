---
title: Trabalhar com datas usando a EXCEL JavaScript
description: Use o Moment-MSDate plug-in com a API javaScript Excel para trabalhar com datas.
ms.date: 02/16/2022
ms.prod: excel
ms.localizationpriority: medium
ms.openlocfilehash: becbbc9deb6f07e244ed0aac1f04b3dad1a800eb
ms.sourcegitcommit: 7b6ee73fa70b8e0ff45c68675dd26dd7a7b8c3e9
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 03/08/2022
ms.locfileid: "63340565"
---
# <a name="work-with-dates-using-the-excel-javascript-api-and-the-moment-msdate-plug-in"></a>Trabalhar com datas usando Excel API JavaScript e o Moment-MSDate plug-in

Este artigo fornece exemplos de código que mostram como trabalhar com datas usando a API JavaScript Excel e o [plug-in Moment-MSDate](https://www.npmjs.com/package/moment-msdate). Para ver a lista completa de propriedades e métodos `Range` compatíveis com o objeto, consulte o Excel[. Classe Range](/javascript/api/excel/excel.range).

[!include[Excel cells and ranges note](../includes/note-excel-cells-and-ranges.md)]

## <a name="use-the-moment-msdate-plug-in-to-work-with-dates"></a>Usar o Moment-MSDate plug-in para trabalhar com datas

A [biblioteca Moment do JavaScript](https://momentjs.com/) fornece uma maneira conveniente de usar datas e carimbos de data e hora. O [plug-in Moment-MSDate](https://www.npmjs.com/package/moment-msdate) converte o formato de momentos em um formato mais apropriado para o Excel. Este é o mesmo formato que a [função NOW](https://support.microsoft.com/office/3337fd29-145a-4347-b2e6-20c904739c46) retorna.

O código a seguir mostra como definir o intervalo em **B4** como um timestamp de um momento.

```js
await Excel.run(async (context) => {
    let sheet = context.workbook.worksheets.getItem("Sample");

    let now = Date.now();
    let nowMoment = moment(now);
    let nowMS = nowMoment.toOADate();

    let dateRange = sheet.getRange("B4");
    dateRange.values = [[nowMS]];

    dateRange.numberFormat = [["[$-409]m/d/yy h:mm AM/PM;@"]];

    await context.sync();
});
```

O exemplo de código a seguir demonstra uma técnica semelhante para obter a data de volta da célula e convertê-la em um `Moment` ou outro formato.

```js
await Excel.run(async (context) => {
    let sheet = context.workbook.worksheets.getItem("Sample");

    let dateRange = sheet.getRange("B4");
    dateRange.load("values");

    await context.sync();

    let nowMS = dateRange.values[0][0];

    // Log the date as a moment.
    let nowMoment = moment.fromOADate(nowMS);
    console.log(`get (moment): ${JSON.stringify(nowMoment)}`);

    // Log the date as a UNIX-style timestamp.
    let now = nowMoment.unix();
    console.log(`get (timestamp): ${now}`);
});
```

Seu complemento precisa formatar os intervalos para exibir as datas em um formulário mais acessível para humanos. Por exemplo, `"[$-409]m/d/yy h:mm AM/PM;@"` exibe "12/3/18 15:57 PM". Para obter mais informações sobre formatos de número de data e hora, consulte "Diretrizes para formatos de data e hora" no artigo Revisar diretrizes para [personalizar um formato de](https://support.microsoft.com/office/c0a1d1fa-d3f4-4018-96b7-9c9354dd99f5) número.

## <a name="see-also"></a>Confira também

- [Trabalhar com células usando a EXCEL JavaScript](excel-add-ins-cells.md)
- [Modelo de objeto JavaScript do Excel em Suplementos do Office](excel-add-ins-core-concepts.md)
- [Trabalhar simultaneamente com vários intervalos em suplementos do Excel](excel-add-ins-multiple-ranges.md)
