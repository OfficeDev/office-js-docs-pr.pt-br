---
title: Trabalhar com datas usando a API JavaScript do Excel
description: Use o Moment-MSDate plug-in com a API JavaScript do Excel para trabalhar com datas.
ms.date: 04/02/2021
ms.prod: excel
localization_priority: Normal
ms.openlocfilehash: d3f59e5daad042541bd933fb4e644d40f27a6e5e
ms.sourcegitcommit: 54fef33bfc7d18a35b3159310bbd8b1c8312f845
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 04/09/2021
ms.locfileid: "51652772"
---
# <a name="work-with-dates-using-the-excel-javascript-api-and-the-moment-msdate-plug-in"></a><span data-ttu-id="dd4e1-103">Trabalhar com datas usando a API JavaScript do Excel e o Moment-MSDate plug-in</span><span class="sxs-lookup"><span data-stu-id="dd4e1-103">Work with dates using the Excel JavaScript API and the Moment-MSDate plug-in</span></span>

<span data-ttu-id="dd4e1-104">Este artigo fornece exemplos de código que mostram como trabalhar com datas usando a API JavaScript do Excel e o [plug-in Moment-MSDate.](https://www.npmjs.com/package/moment-msdate)</span><span class="sxs-lookup"><span data-stu-id="dd4e1-104">This article provides code samples that show how to work with dates using the Excel JavaScript API and the [Moment-MSDate plug-in](https://www.npmjs.com/package/moment-msdate).</span></span> <span data-ttu-id="dd4e1-105">Para ver a lista completa de propriedades e métodos compatíveis com o `Range` objeto, consulte a [classe Excel.Range](/javascript/api/excel/excel.range).</span><span class="sxs-lookup"><span data-stu-id="dd4e1-105">For the complete list of properties and methods that the `Range` object supports, see the [Excel.Range class](/javascript/api/excel/excel.range).</span></span>

[!include[Excel cells and ranges note](../includes/note-excel-cells-and-ranges.md)]

## <a name="use-the-moment-msdate-plug-in-to-work-with-dates"></a><span data-ttu-id="dd4e1-106">Usar o Moment-MSDate plug-in para trabalhar com datas</span><span class="sxs-lookup"><span data-stu-id="dd4e1-106">Use the Moment-MSDate plug-in to work with dates</span></span>

<span data-ttu-id="dd4e1-107">A [biblioteca Moment do JavaScript](https://momentjs.com/) fornece uma maneira conveniente de usar datas e carimbos de data e hora.</span><span class="sxs-lookup"><span data-stu-id="dd4e1-107">The [Moment JavaScript library](https://momentjs.com/) provides a convenient way to use dates and timestamps.</span></span> <span data-ttu-id="dd4e1-108">O [plug-in Moment-MSDate](https://www.npmjs.com/package/moment-msdate) converte o formato de momentos em um formato mais apropriado para o Excel.</span><span class="sxs-lookup"><span data-stu-id="dd4e1-108">The [Moment-MSDate plug-in](https://www.npmjs.com/package/moment-msdate) converts the format of moments into one preferable for Excel.</span></span> <span data-ttu-id="dd4e1-109">Este é o mesmo formato que a [função NOW](https://support.office.com/article/now-function-3337fd29-145a-4347-b2e6-20c904739c46) retorna.</span><span class="sxs-lookup"><span data-stu-id="dd4e1-109">This is the same format the [NOW function](https://support.office.com/article/now-function-3337fd29-145a-4347-b2e6-20c904739c46) returns.</span></span>

<span data-ttu-id="dd4e1-110">O código a seguir mostra como definir o intervalo em **B4** como um timestamp de um momento.</span><span class="sxs-lookup"><span data-stu-id="dd4e1-110">The following code shows how to set the range at **B4** to a moment's timestamp.</span></span>

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

<span data-ttu-id="dd4e1-111">O exemplo de código a seguir demonstra uma técnica semelhante para obter a data de volta da célula e convertê-la em `Moment` um ou outro formato.</span><span class="sxs-lookup"><span data-stu-id="dd4e1-111">The following code sample demonstrates a similar technique to get the date back out of the cell and convert it to a `Moment` or other format.</span></span>

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

<span data-ttu-id="dd4e1-112">Seu complemento precisa formatar os intervalos para exibir as datas em um formulário mais acessível para humanos.</span><span class="sxs-lookup"><span data-stu-id="dd4e1-112">Your add-in has to format the ranges to display the dates in a more human-readable form.</span></span> <span data-ttu-id="dd4e1-113">Por exemplo, `"[$-409]m/d/yy h:mm AM/PM;@"` exibe "12/3/18 15:57 PM".</span><span class="sxs-lookup"><span data-stu-id="dd4e1-113">For example, `"[$-409]m/d/yy h:mm AM/PM;@"` displays "12/3/18 3:57 PM".</span></span> <span data-ttu-id="dd4e1-114">Para obter mais informações sobre formatos de número de data e hora, consulte "Diretrizes para formatos de data e hora" no artigo Revisar diretrizes para personalizar um [formato de](https://support.microsoft.com/office/c0a1d1fa-d3f4-4018-96b7-9c9354dd99f5) número.</span><span class="sxs-lookup"><span data-stu-id="dd4e1-114">For more information about date and time number formats, see "Guidelines for date and time formats" in the [Review guidelines for customizing a number format](https://support.microsoft.com/office/c0a1d1fa-d3f4-4018-96b7-9c9354dd99f5) article.</span></span>


## <a name="see-also"></a><span data-ttu-id="dd4e1-115">Confira também</span><span class="sxs-lookup"><span data-stu-id="dd4e1-115">See also</span></span>

- [<span data-ttu-id="dd4e1-116">Trabalhar com células usando a API JavaScript do Excel</span><span class="sxs-lookup"><span data-stu-id="dd4e1-116">Work with cells using the Excel JavaScript API</span></span>](excel-add-ins-cells.md)
- [<span data-ttu-id="dd4e1-117">Modelo de objeto JavaScript do Excel em Suplementos do Office</span><span class="sxs-lookup"><span data-stu-id="dd4e1-117">Excel JavaScript object model in Office Add-ins</span></span>](excel-add-ins-core-concepts.md)
- [<span data-ttu-id="dd4e1-118">Trabalhar simultaneamente com vários intervalos em suplementos do Excel</span><span class="sxs-lookup"><span data-stu-id="dd4e1-118">Work with multiple ranges simultaneously in Excel add-ins</span></span>](excel-add-ins-multiple-ranges.md)
