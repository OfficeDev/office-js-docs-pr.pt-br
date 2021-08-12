---
title: Trabalhar com datas usando a EXCEL JavaScript
description: Use o Moment-MSDate plug-in com a API javaScript Excel para trabalhar com datas.
ms.date: 04/02/2021
ms.prod: excel
localization_priority: Normal
ms.openlocfilehash: fdfc39f12b3374d9903156b1ba71a9bbd4f296735f0ed41dac56d62243058c1d
ms.sourcegitcommit: 4f2c76b48d15e7d03c5c5f1f809493758fcd88ec
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 08/07/2021
ms.locfileid: "57084726"
---
# <a name="work-with-dates-using-the-excel-javascript-api-and-the-moment-msdate-plug-in"></a>Trabalhar com datas usando Excel API JavaScript e o Moment-MSDate plug-in

Este artigo fornece exemplos de código que mostram como trabalhar com datas usando a API JavaScript Excel e o [plug-in Moment-MSDate.](https://www.npmjs.com/package/moment-msdate) Para ver a lista completa de propriedades e métodos compatíveis com o `Range` objeto, consulte o [Excel. Classe Range](/javascript/api/excel/excel.range).

[!include[Excel cells and ranges note](../includes/note-excel-cells-and-ranges.md)]

## <a name="use-the-moment-msdate-plug-in-to-work-with-dates"></a>Usar o Moment-MSDate plug-in para trabalhar com datas

A [biblioteca Moment do JavaScript](https://momentjs.com/) fornece uma maneira conveniente de usar datas e carimbos de data e hora. O [plug-in Moment-MSDate](https://www.npmjs.com/package/moment-msdate) converte o formato de momentos em um formato mais apropriado para o Excel. Este é o mesmo formato que a [função NOW](https://support.office.com/article/now-function-3337fd29-145a-4347-b2e6-20c904739c46) retorna.

O código a seguir mostra como definir o intervalo em **B4** como um timestamp de um momento.

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

O exemplo de código a seguir demonstra uma técnica semelhante para obter a data de volta da célula e convertê-la em `Moment` um ou outro formato.

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

Seu complemento precisa formatar os intervalos para exibir as datas em um formulário mais acessível para humanos. Por exemplo, `"[$-409]m/d/yy h:mm AM/PM;@"` exibe "12/3/18 15:57 PM". Para obter mais informações sobre formatos de número de data e hora, consulte "Diretrizes para formatos de data e hora" no artigo Revisar diretrizes para personalizar um [formato de](https://support.microsoft.com/office/c0a1d1fa-d3f4-4018-96b7-9c9354dd99f5) número.


## <a name="see-also"></a>Confira também

- [Trabalhar com células usando a EXCEL JavaScript](excel-add-ins-cells.md)
- [Modelo de objeto JavaScript do Excel em Suplementos do Office](excel-add-ins-core-concepts.md)
- [Trabalhar simultaneamente com vários intervalos em suplementos do Excel](excel-add-ins-multiple-ranges.md)
