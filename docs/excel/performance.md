---
title: Otimização de desempenho do da API JavaScript do Excel
description: Otimizar o desempenho do suplemento do Excel usando a API JavaScript.
ms.date: 07/29/2020
localization_priority: Normal
ms.openlocfilehash: 42ab5f28717f0f7dcd06461840de692a5daf60ce
ms.sourcegitcommit: c6308cf245ac1bc66a876eaa0a7bb4a2492991ac
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 09/08/2020
ms.locfileid: "47408611"
---
# <a name="performance-optimization-using-the-excel-javascript-api"></a><span data-ttu-id="91796-103">Otimização de desempenho usando a API JavaScript do Excel</span><span class="sxs-lookup"><span data-stu-id="91796-103">Performance optimization using the Excel JavaScript API</span></span>

<span data-ttu-id="91796-104">Existem várias maneiras de executar tarefas comuns com a API JavaScript do Excel.</span><span class="sxs-lookup"><span data-stu-id="91796-104">There are multiple ways that you can perform common tasks with the Excel JavaScript API.</span></span> <span data-ttu-id="91796-105">Você encontrará diferenças significativas de desempenho entre várias abordagens.</span><span class="sxs-lookup"><span data-stu-id="91796-105">You'll find significant performance differences between various approaches.</span></span> <span data-ttu-id="91796-106">Este artigo fornece orientações e amostras de código para mostrar como realizar tarefas comuns com eficiência usando as API JavaScript do Excel.</span><span class="sxs-lookup"><span data-stu-id="91796-106">This article provides guidance and code samples to show you how to perform common tasks efficiently using Excel JavaScript API.</span></span>

> [!IMPORTANT]
> <span data-ttu-id="91796-107">Muitos problemas de desempenho podem ser tratados através do uso recomendado de `load` `sync` chamadas e.</span><span class="sxs-lookup"><span data-stu-id="91796-107">Many performance issues can be addressed through recommended usage of `load` and `sync` calls.</span></span> <span data-ttu-id="91796-108">Consulte a seção "aprimoramentos de desempenho com as APIs específicas do aplicativo" de [limites de recurso e otimização de desempenho para suplementos do Office](../concepts/resource-limits-and-performance-optimization.md#performance-improvements-with-the-application-specific-apis) para conselhos sobre como trabalhar com APIs específicas do aplicativo de uma maneira eficiente.</span><span class="sxs-lookup"><span data-stu-id="91796-108">See the "Performance improvements with the application-specific APIs" section of [Resource limits and performance optimization for Office Add-ins](../concepts/resource-limits-and-performance-optimization.md#performance-improvements-with-the-application-specific-apis) for advice on working with the application-specific APIs in an efficient way.</span></span>

## <a name="suspend-excel-processes-temporarily"></a><span data-ttu-id="91796-109">Suspender temporariamente os processos do Excel</span><span class="sxs-lookup"><span data-stu-id="91796-109">Suspend Excel processes temporarily</span></span>

<span data-ttu-id="91796-110">O Excel tem várias tarefas em segundo plano reagindo à entrada de usuários e seu suplemento.</span><span class="sxs-lookup"><span data-stu-id="91796-110">Excel has a number of background tasks reacting to input from both users and your add-in.</span></span> <span data-ttu-id="91796-111">Alguns desses processos do Excel podem ser controlado para obter o benefício de desempenho.</span><span class="sxs-lookup"><span data-stu-id="91796-111">Some of these Excel processes can be controlled to yield a performance benefit.</span></span> <span data-ttu-id="91796-112">Isso é útil principalmente quando o suplemento lida com grandes conjuntos de dados.</span><span class="sxs-lookup"><span data-stu-id="91796-112">This is especially helpful when your add-in deals with large data sets.</span></span>

### <a name="suspend-calculation-temporarily"></a><span data-ttu-id="91796-113">Suspender os cálculos temporariamente</span><span class="sxs-lookup"><span data-stu-id="91796-113">Suspend calculation temporarily</span></span>

<span data-ttu-id="91796-114">Se você estiver tentando executar uma operação em um grande número de células (por exemplo, definindo o valor do objeto de um grande intervalo) e não se importar em suspender o cálculo no Excel temporariamente enquanto a operação for concluída, é recomendável que você suspenda o cálculo até o próximo `context.sync()` ser chamado.</span><span class="sxs-lookup"><span data-stu-id="91796-114">If you are trying to perform an operation on a large number of cells (for example, setting the value of a huge range object) and you don't mind suspending the calculation in Excel temporarily while your operation finishes, we recommend that you suspend calculation until the next `context.sync()` is called.</span></span>

<span data-ttu-id="91796-115">Ver a documentação de referência [objeto de aplicativo](/javascript/api/excel/excel.application) para saber mais sobre como usar a API`suspendApiCalculationUntilNextSync()`para suspender e reativar cálculos de maneira muito fácil.</span><span class="sxs-lookup"><span data-stu-id="91796-115">See the [Application Object](/javascript/api/excel/excel.application) reference documentation for information about how to use the `suspendApiCalculationUntilNextSync()` API to suspend and reactivate calculations in a very convenient way.</span></span> <span data-ttu-id="91796-116">O código a seguir demonstra como suspender temporariamente um cálculo:</span><span class="sxs-lookup"><span data-stu-id="91796-116">The following code demonstrates how to suspend calculation temporarily:</span></span>

```js
Excel.run(async function(ctx) {
    var app = ctx.workbook.application;
    var sheet = ctx.workbook.worksheets.getItem("sheet1");
    var rangeToSet: Excel.Range;
    var rangeToGet: Excel.Range;
    app.load("calculationMode");
    await ctx.sync();
    // Calculation mode should be "Automatic" by default
    console.log(app.calculationMode);

    rangeToSet = sheet.getRange("A1:C1");
    rangeToSet.values = [[1, 2, "=SUM(A1:B1)"]];
    rangeToGet = sheet.getRange("A1:C1");
    rangeToGet.load("values");
    await ctx.sync();
    // Range value should be [1, 2, 3] now
    console.log(rangeToGet.values);

    // Suspending recalculation
    app.suspendApiCalculationUntilNextSync();
    rangeToSet = sheet.getRange("A1:B1");
    rangeToSet.values = [[10, 20]];
    rangeToGet = sheet.getRange("A1:C1");
    rangeToGet.load("values");
    app.load("calculationMode");
    await ctx.sync();
    // Range value should be [10, 20, 3] when we load the property, because calculation is suspended at that point
    console.log(rangeToGet.values);
    // Calculation mode should still be "Automatic" even with suspend recalculation
    console.log(app.calculationMode);

    rangeToGet.load("values");
    await ctx.sync();
    // Range value should be [10, 20, 30] when we load the property, because calculation is resumed after last sync
    console.log(rangeToGet.values);
})
```

<span data-ttu-id="91796-117">Observe que somente os cálculos de fórmula são suspensos.</span><span class="sxs-lookup"><span data-stu-id="91796-117">Please note that only formula calculations are suspended.</span></span> <span data-ttu-id="91796-118">Todas as referências alteradas ainda serão recriadas.</span><span class="sxs-lookup"><span data-stu-id="91796-118">Any altered references are still rebuilt.</span></span> <span data-ttu-id="91796-119">Por exemplo, renomear uma planilha ainda atualiza quaisquer referências em fórmulas para essa planilha.</span><span class="sxs-lookup"><span data-stu-id="91796-119">For example, renaming a worksheet still updates any references in formulas to that worksheet.</span></span>

### <a name="suspend-screen-updating"></a><span data-ttu-id="91796-120">Suspender a atualização da tela</span><span class="sxs-lookup"><span data-stu-id="91796-120">Suspend screen updating</span></span>

<span data-ttu-id="91796-121">O Excel exibe as alterações que seu suplemento faz aproximadamente conforme elas acontecem no código.</span><span class="sxs-lookup"><span data-stu-id="91796-121">Excel displays changes your add-in makes approximately as they happen in the code.</span></span> <span data-ttu-id="91796-122">Para conjuntos de dados grandes e interativos, talvez não seja necessário não esse andamento na tela em tempo real.</span><span class="sxs-lookup"><span data-stu-id="91796-122">For large, iterative data sets, you may not need to see this progress on the screen in real-time.</span></span> <span data-ttu-id="91796-123">`Application.suspendScreenUpdatingUntilNextSync()` pausa atualizações visuais no Excel até as chamadas do suplemento `context.sync()`, ou até o`Excel.run` terminar (chamadas implícitas `context.sync`).</span><span class="sxs-lookup"><span data-stu-id="91796-123">`Application.suspendScreenUpdatingUntilNextSync()` pauses visual updates to Excel until the add-in calls `context.sync()`, or until `Excel.run` ends (implicitly calling `context.sync`).</span></span> <span data-ttu-id="91796-124">Lembre-se, o Excel não mostrará os sinais de atividade até a próxima sincronização. Seu suplemento deve fornecer orientação aos usuários para prepará-los para esse atraso ou fornecer uma barra de status para demonstrar atividade.</span><span class="sxs-lookup"><span data-stu-id="91796-124">Be aware, Excel will not show any signs of activity until the next sync. Your add-in should either give users guidance to prepare them for this delay or provide a status bar to demonstrate activity.</span></span>

> [!NOTE]
> <span data-ttu-id="91796-125">Não chame `suspendScreenUpdatingUntilNextSync` repetidamente (como em um loop).</span><span class="sxs-lookup"><span data-stu-id="91796-125">Don't call `suspendScreenUpdatingUntilNextSync` repeatedly (such as in a loop).</span></span> <span data-ttu-id="91796-126">As chamadas repetidas farão com que a janela do Excel fique de piscar.</span><span class="sxs-lookup"><span data-stu-id="91796-126">Repeated calls will cause the Excel window to flicker.</span></span>

### <a name="enable-and-disable-events"></a><span data-ttu-id="91796-127">Habilitar e desabilitar eventos</span><span class="sxs-lookup"><span data-stu-id="91796-127">Enable and disable events</span></span>

<span data-ttu-id="91796-128">O desempenho de um suplemento pode ser melhorado desabilitando eventos.</span><span class="sxs-lookup"><span data-stu-id="91796-128">Performance of an add-in may be improved by disabling events.</span></span> <span data-ttu-id="91796-129">Um exemplo de código mostrando como habilitar e desabilitar os eventos está no artigo [trabalhar com eventos](excel-add-ins-events.md#enable-and-disable-events).</span><span class="sxs-lookup"><span data-stu-id="91796-129">A code sample showing how to enable and disable events is in the [Work with Events](excel-add-ins-events.md#enable-and-disable-events) article.</span></span>

## <a name="importing-data-into-tables"></a><span data-ttu-id="91796-130">Importar dados em tabelas</span><span class="sxs-lookup"><span data-stu-id="91796-130">Importing data into tables</span></span>

<span data-ttu-id="91796-131">Ao tentar importar um grande volume de dados diretamente em um objeto[tabela](/javascript/api/excel/excel.table) diretamente (por exemplo, usando `TableRowCollection.add()`), você poderá observar um desempenho lento.</span><span class="sxs-lookup"><span data-stu-id="91796-131">When trying to import a huge amount of data directly into a [Table](/javascript/api/excel/excel.table) object directly (for example, by using `TableRowCollection.add()`), you might experience slow performance.</span></span> <span data-ttu-id="91796-132">Se você estiver tentando adicionar uma nova tabela, você deve preencher os dados primeiro definindo `range.values`e em seguida, ligue `worksheet.tables.add()` para criar uma tabela de intervalo.</span><span class="sxs-lookup"><span data-stu-id="91796-132">If you are trying to add a new table, you should fill in the data first by setting `range.values`, and then call `worksheet.tables.add()` to create a table over the range.</span></span> <span data-ttu-id="91796-133">Se você está tentando gravar dados em uma tabela existente, grave os dados em um intervalo de objeto via`table.getDataBodyRange()`, e a tabela será expandida automaticamente.</span><span class="sxs-lookup"><span data-stu-id="91796-133">If you are trying to write data into an existing table, write the data into a range object via `table.getDataBodyRange()`, and the table will expand automatically.</span></span>

<span data-ttu-id="91796-134">Aqui está um exemplo dessa abordagem:</span><span class="sxs-lookup"><span data-stu-id="91796-134">Here is an example of this approach:</span></span>

```js
Excel.run(async (ctx) => {
    var sheet = ctx.workbook.worksheets.getItem("Sheet1");
    // Write the data into the range first.
    var range = sheet.getRange("A1:B3");
    range.values = [["Key", "Value"], ["A", 1], ["B", 2]];

    // Create the table over the range
    var table = sheet.tables.add('A1:B3', true);
    table.name = "Example";
    await ctx.sync();


    // Insert a new row to the table
    table.getDataBodyRange().getRowsBelow(1).values = [["C", 3]];
    // Change a existing row value
    table.getDataBodyRange().getRow(1).values = [["D", 4]];
    await ctx.sync();
})
```

> [!NOTE]
> <span data-ttu-id="91796-135">Você pode converter convenientemente um objeto de tabela em um objeto de intervalo usando o método[Table.convertToRange()](/javascript/api/excel/excel.table#converttorange--).</span><span class="sxs-lookup"><span data-stu-id="91796-135">You can conveniently convert a Table object to a Range object by using the [Table.convertToRange()](/javascript/api/excel/excel.table#converttorange--) method.</span></span>

## <a name="see-also"></a><span data-ttu-id="91796-136">Confira também</span><span class="sxs-lookup"><span data-stu-id="91796-136">See also</span></span>

* [<span data-ttu-id="91796-137">Modelo de objeto do JavaScript do Excel em suplementos do Office</span><span class="sxs-lookup"><span data-stu-id="91796-137">Excel JavaScript object model in Office Add-ins</span></span>](excel-add-ins-core-concepts.md)
* [<span data-ttu-id="91796-138">Limites de recurso e otimização de desempenho para Suplementos do Office</span><span class="sxs-lookup"><span data-stu-id="91796-138">Resource limits and performance optimization for Office Add-ins</span></span>](../concepts/resource-limits-and-performance-optimization.md)
* [<span data-ttu-id="91796-139">Objeto de funções de planilha (API JavaScript para Excel)</span><span class="sxs-lookup"><span data-stu-id="91796-139">Worksheet Functions Object (JavaScript API for Excel)</span></span>](/javascript/api/excel/excel.functions)
