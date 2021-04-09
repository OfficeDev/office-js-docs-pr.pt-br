---
title: Obter um intervalo usando a API JavaScript do Excel
description: Saiba como recuperar um intervalo usando a API JavaScript do Excel.
ms.date: 04/02/2021
ms.prod: excel
localization_priority: Normal
ms.openlocfilehash: 6aa9bb00bc9d24aeee5f1fef9e8d1531525e9d1f
ms.sourcegitcommit: 54fef33bfc7d18a35b3159310bbd8b1c8312f845
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 04/09/2021
ms.locfileid: "51652770"
---
# <a name="get-a-range-using-the-excel-javascript-api"></a><span data-ttu-id="e2dd9-103">Obter um intervalo usando a API JavaScript do Excel</span><span class="sxs-lookup"><span data-stu-id="e2dd9-103">Get a range using the Excel JavaScript API</span></span>

<span data-ttu-id="e2dd9-104">Este artigo fornece exemplos que mostram diferentes maneiras de obter um intervalo dentro de uma planilha usando a API JavaScript do Excel.</span><span class="sxs-lookup"><span data-stu-id="e2dd9-104">This article provides examples that show different ways to get a range within a worksheet using the Excel JavaScript API.</span></span> <span data-ttu-id="e2dd9-105">Para ver a lista completa de propriedades e métodos compatíveis com o `Range` objeto, consulte [Classe Excel.Range](/javascript/api/excel/excel.range).</span><span class="sxs-lookup"><span data-stu-id="e2dd9-105">For the complete list of properties and methods that the `Range` object supports, see [Excel.Range class](/javascript/api/excel/excel.range).</span></span>

[!include[Excel cells and ranges note](../includes/note-excel-cells-and-ranges.md)]

## <a name="get-range-by-address"></a><span data-ttu-id="e2dd9-106">Obter intervalo por endereço</span><span class="sxs-lookup"><span data-stu-id="e2dd9-106">Get range by address</span></span>

<span data-ttu-id="e2dd9-107">O exemplo de código a seguir obtém o intervalo com o endereço **B2:C5** da planilha denominada **Exemplo**, carrega sua propriedade e grava uma mensagem `address` no console.</span><span class="sxs-lookup"><span data-stu-id="e2dd9-107">The following code sample gets the range with address **B2:C5** from the worksheet named **Sample**, loads its `address` property, and writes a message to the console.</span></span>

```js
Excel.run(function (context) {
    var sheet = context.workbook.worksheets.getItem("Sample");
    var range = sheet.getRange("B2:C5");
    range.load("address");

    return context.sync()
        .then(function () {
            console.log(`The address of the range B2:C5 is "${range.address}"`);
        });
}).catch(errorHandlerFunction);
```

## <a name="get-range-by-name"></a><span data-ttu-id="e2dd9-108">Obter intervalo por nome</span><span class="sxs-lookup"><span data-stu-id="e2dd9-108">Get range by name</span></span>

<span data-ttu-id="e2dd9-109">O exemplo de código a seguir obtém o intervalo nomeado da planilha denominada Exemplo , carrega sua propriedade e grava `MyRange` uma mensagem no  `address` console.</span><span class="sxs-lookup"><span data-stu-id="e2dd9-109">The following code sample gets the range named `MyRange` from the worksheet named **Sample**, loads its `address` property, and writes a message to the console.</span></span>

```js
Excel.run(function (context) {
    var sheet = context.workbook.worksheets.getItem("Sample");
    var range = sheet.getRange("MyRange");
    range.load("address");

    return context.sync()
        .then(function () {
            console.log(`The address of the range "MyRange" is "${range.address}"`);
        });
}).catch(errorHandlerFunction);
```

## <a name="get-used-range"></a><span data-ttu-id="e2dd9-110">Obter intervalo usado</span><span class="sxs-lookup"><span data-stu-id="e2dd9-110">Get used range</span></span>

<span data-ttu-id="e2dd9-111">O exemplo de código a seguir obtém o intervalo usado da planilha denominada **Exemplo**, carrega sua propriedade e grava `address` uma mensagem no console.</span><span class="sxs-lookup"><span data-stu-id="e2dd9-111">The following code sample gets the used range from the worksheet named **Sample**, loads its `address` property, and writes a message to the console.</span></span> <span data-ttu-id="e2dd9-112">O intervalo usado é o menor intervalo que abrange todas as células na planilha que têm um valor ou uma formatação atribuída a elas.</span><span class="sxs-lookup"><span data-stu-id="e2dd9-112">The used range is the smallest range that encompasses any cells in the worksheet that have a value or formatting assigned to them.</span></span> <span data-ttu-id="e2dd9-113">Se a planilha inteira estiver em branco, o método retornará um intervalo que `getUsedRange()` consiste apenas na célula superior esquerda.</span><span class="sxs-lookup"><span data-stu-id="e2dd9-113">If the entire worksheet is blank, the `getUsedRange()` method returns a range that consists of only the top-left cell.</span></span>

```js
Excel.run(function (context) {
    var sheet = context.workbook.worksheets.getItem("Sample");
    var range = sheet.getUsedRange();
    range.load("address");

    return context.sync()
        .then(function () {
            console.log(`The address of the used range in the worksheet is "${range.address}"`);
        });
}).catch(errorHandlerFunction);
```

## <a name="get-entire-range"></a><span data-ttu-id="e2dd9-114">Obter intervalo inteiro</span><span class="sxs-lookup"><span data-stu-id="e2dd9-114">Get entire range</span></span>

<span data-ttu-id="e2dd9-115">O exemplo de código a seguir obtém todo o intervalo de planilhas da planilha denominada **Exemplo**, carrega sua propriedade e grava `address` uma mensagem no console.</span><span class="sxs-lookup"><span data-stu-id="e2dd9-115">The following code sample gets the entire worksheet range from the worksheet named **Sample**, loads its `address` property, and writes a message to the console.</span></span>

```js
Excel.run(function (context) {
    var sheet = context.workbook.worksheets.getItem("Sample");
    var range = sheet.getRange();
    range.load("address");

    return context.sync()
        .then(function () {
            console.log(`The address of the entire worksheet range is "${range.address}"`);
        });
}).catch(errorHandlerFunction);
```

## <a name="see-also"></a><span data-ttu-id="e2dd9-116">Confira também</span><span class="sxs-lookup"><span data-stu-id="e2dd9-116">See also</span></span>

- [<span data-ttu-id="e2dd9-117">Modelo de objeto JavaScript do Excel em Suplementos do Office</span><span class="sxs-lookup"><span data-stu-id="e2dd9-117">Excel JavaScript object model in Office Add-ins</span></span>](excel-add-ins-core-concepts.md)
- [<span data-ttu-id="e2dd9-118">Trabalhar com células usando a API JavaScript do Excel</span><span class="sxs-lookup"><span data-stu-id="e2dd9-118">Work with cells using the Excel JavaScript API</span></span>](excel-add-ins-cells.md)
- [<span data-ttu-id="e2dd9-119">Inserir um intervalo usando a API JavaScript do Excel</span><span class="sxs-lookup"><span data-stu-id="e2dd9-119">Insert a range using the Excel JavaScript API</span></span>](excel-add-ins-ranges-insert.md)
