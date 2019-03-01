---
title: Trabalhar com intervalos usando a API JavaScript do Excel (fundamental)
description: ''
ms.date: 02/20/2019
localization_priority: Priority
ms.openlocfilehash: c911bf6e908d06298e768bdaf047262d78808321
ms.sourcegitcommit: 8e20e7663be2aaa0f7a5436a965324d171bc667d
ms.translationtype: HT
ms.contentlocale: pt-BR
ms.lasthandoff: 02/28/2019
ms.locfileid: "30199582"
---
# <a name="work-with-ranges-using-the-excel-javascript-api"></a><span data-ttu-id="47169-102">Trabalhar com intervalos usando a API JavaScript do Excel</span><span class="sxs-lookup"><span data-stu-id="47169-102">Work with ranges using the Excel JavaScript API</span></span>

<span data-ttu-id="47169-103">Este artigo fornece exemplos de código que mostram como executar tarefas comuns com intervalos usando a API JavaScript do Excel.</span><span class="sxs-lookup"><span data-stu-id="47169-103">This article provides code samples that show how to perform common tasks with ranges using the Excel JavaScript API.</span></span> <span data-ttu-id="47169-104">Para obter a lista completa de propriedades e métodos que o objeto **Range** suporta, confira [Objeto Range (API JavaScript para Excel)](/javascript/api/excel/excel.range).</span><span class="sxs-lookup"><span data-stu-id="47169-104">For the complete list of properties and methods that the **Range** object supports, see [Range Object (JavaScript API for Excel)](/javascript/api/excel/excel.range).</span></span>

> [!NOTE]
> <span data-ttu-id="47169-105">Confira exemplos de código que mostram como executar tarefas avançadas com intervalos em [Trabalhar com intervalos usando a API JavaScript do Excel (avançado)](excel-add-ins-ranges-advanced.md).</span><span class="sxs-lookup"><span data-stu-id="47169-105">For code samples that show how to perform more advanced tasks with ranges, see [Work with ranges using the Excel JavaScript API (advanced)](excel-add-ins-ranges-advanced.md).</span></span>

## <a name="get-a-range"></a><span data-ttu-id="47169-106">Obter um intervalo</span><span class="sxs-lookup"><span data-stu-id="47169-106">Get a range</span></span>

<span data-ttu-id="47169-107">Os exemplos a seguir mostram diferentes maneiras de obter uma referência a um intervalo em uma planilha.</span><span class="sxs-lookup"><span data-stu-id="47169-107">The following examples show different ways to get a reference to a range within a worksheet.</span></span>

### <a name="get-range-by-address"></a><span data-ttu-id="47169-108">Obter intervalo por endereço</span><span class="sxs-lookup"><span data-stu-id="47169-108">Get range by address</span></span>

<span data-ttu-id="47169-109">O exemplo de código a seguir obtém o intervalo com o endereço **B2:B5** da planilha chamada **Amostra**, carrega sua propriedade **address** e grava uma mensagem no console.</span><span class="sxs-lookup"><span data-stu-id="47169-109">The following code sample gets the range with address **B2:B5** from the worksheet named **Sample**, loads its **address** property, and writes a message to the console.</span></span>

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

### <a name="get-range-by-name"></a><span data-ttu-id="47169-110">Obter intervalo por nome</span><span class="sxs-lookup"><span data-stu-id="47169-110">Get range by name</span></span>

<span data-ttu-id="47169-111">O exemplo de código a seguir obtém o intervalo chamado **MyRange** da planilha chamada **Amostra**, carrega sua propriedade **address** e grava uma mensagem no console.</span><span class="sxs-lookup"><span data-stu-id="47169-111">The following code sample gets the range named **MyRange** from the worksheet named **Sample**, loads its **address** property, and writes a message to the console.</span></span>

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

### <a name="get-used-range"></a><span data-ttu-id="47169-112">Obter intervalo usado</span><span class="sxs-lookup"><span data-stu-id="47169-112">Get used range</span></span>

<span data-ttu-id="47169-113">O exemplo de código a seguir obtém o intervalo usado da planilha chamada **Amostra**, carrega sua propriedade **address** e grava uma mensagem no console.</span><span class="sxs-lookup"><span data-stu-id="47169-113">The following code sample gets the used range from the worksheet named **Sample**, loads its **address** property, and writes a message to the console.</span></span> <span data-ttu-id="47169-114">O intervalo usado é o menor intervalo que abrange todas as células na planilha que têm um valor ou uma formatação atribuída a elas.</span><span class="sxs-lookup"><span data-stu-id="47169-114">The used range is the smallest range that encompasses any cells in the worksheet that have a value or formatting assigned to them.</span></span> <span data-ttu-id="47169-115">Se a planilha inteira estiver em branco, o método **getUsedRange()** retornará um intervalo que consiste apenas na célula superior esquerda da planilha.</span><span class="sxs-lookup"><span data-stu-id="47169-115">If the entire worksheet is blank, the **getUsedRange()** method returns a range that consists of only the top-left cell in the worksheet.</span></span>

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

### <a name="get-entire-range"></a><span data-ttu-id="47169-116">Obter intervalo inteiro</span><span class="sxs-lookup"><span data-stu-id="47169-116">Get entire range</span></span>

<span data-ttu-id="47169-117">O exemplo de código a seguir obtém todo o intervalo da planilha chamada **Amostra**, carrega sua propriedade **address** e grava uma mensagem no console.</span><span class="sxs-lookup"><span data-stu-id="47169-117">The following code sample gets the entire worksheet range from the worksheet named **Sample**, loads its **address** property, and writes a message to the console.</span></span>

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

## <a name="insert-a-range-of-cells"></a><span data-ttu-id="47169-118">Inserir um intervalo de células</span><span class="sxs-lookup"><span data-stu-id="47169-118">Insert a range of cells</span></span>

<span data-ttu-id="47169-119">O exemplo de código a seguir insere um intervalo de células no local **B4:E4** e desloca outras células para baixo a fim de fornecer espaço para as novas células.</span><span class="sxs-lookup"><span data-stu-id="47169-119">The following code sample inserts a range of cells in location **B4:E4** and shifts other cells down to provide space for the new cells.</span></span>

```js
Excel.run(function (context) {
    var sheet = context.workbook.worksheets.getItem("Sample");
    var range = sheet.getRange("B4:E4");

    range.insert(Excel.InsertShiftDirection.down);
    
    return context.sync();
}).catch(errorHandlerFunction);
```

<span data-ttu-id="47169-120">**Dados antes da inserção do intervalo**</span><span class="sxs-lookup"><span data-stu-id="47169-120">**Data before range is inserted**</span></span>

![Dados no Excel antes da inserção do intervalo](../images/excel-ranges-start.png)

<span data-ttu-id="47169-122">**Dados após a inserção do intervalo**</span><span class="sxs-lookup"><span data-stu-id="47169-122">**Data after range is inserted**</span></span>

![Dados no Excel após a inserção do intervalo](../images/excel-ranges-after-insert.png)

## <a name="clear-a-range-of-cells"></a><span data-ttu-id="47169-124">Limpar um intervalo de células</span><span class="sxs-lookup"><span data-stu-id="47169-124">Clear a range of cells</span></span>

<span data-ttu-id="47169-125">O exemplo de código a seguir limpa todo o conteúdo e a formatação das células no intervalo **E2:E5**.</span><span class="sxs-lookup"><span data-stu-id="47169-125">The following code sample clears all contents and formatting of cells in the range **E2:E5**.</span></span>  

```js
Excel.run(function (context) {
    var sheet = context.workbook.worksheets.getItem("Sample");
    var range = sheet.getRange("E2:E5");

    range.clear();

    return context.sync();
}).catch(errorHandlerFunction);
```

<span data-ttu-id="47169-126">**Dados antes da limpeza do intervalo**</span><span class="sxs-lookup"><span data-stu-id="47169-126">**Data before range is cleared**</span></span>

![Dados no Excel antes da limpeza do intervalo](../images/excel-ranges-start.png)

<span data-ttu-id="47169-128">**Dados após a limpeza do intervalo**</span><span class="sxs-lookup"><span data-stu-id="47169-128">**Data after range is cleared**</span></span>

![Dados no Excel após a limpeza do intervalo](../images/excel-ranges-after-clear.png)

## <a name="delete-a-range-of-cells"></a><span data-ttu-id="47169-130">Excluir um intervalo de células</span><span class="sxs-lookup"><span data-stu-id="47169-130">Delete a range of cells</span></span>

<span data-ttu-id="47169-131">O exemplo de código a seguir exclui as células no intervalo **B4:E4** e desloca outras células para cima a fim de preencher o espaço deixado pelas células excluídas.</span><span class="sxs-lookup"><span data-stu-id="47169-131">The following code sample deletes the cells in the range **B4:E4** and shift other cells up to fill the space that was vacated by the deleted cells.</span></span>

```js
Excel.run(function (context) {
    var sheet = context.workbook.worksheets.getItem("Sample");
    var range = sheet.getRange("B4:E4");

    range.delete(Excel.DeleteShiftDirection.up);

    return context.sync();
}).catch(errorHandlerFunction);
```

<span data-ttu-id="47169-132">**Dados antes da exclusão do intervalo**</span><span class="sxs-lookup"><span data-stu-id="47169-132">**Data before range is deleted**</span></span>

![Dados no Excel antes da exclusão do intervalo](../images/excel-ranges-start.png)

<span data-ttu-id="47169-134">**Dados após a exclusão do intervalo**</span><span class="sxs-lookup"><span data-stu-id="47169-134">**Data after range is deleted**</span></span>

![Dados no Excel após a exclusão do intervalo](../images/excel-ranges-after-delete.png)

## <a name="set-the-selected-range"></a><span data-ttu-id="47169-136">Definir o intervalo selecionado</span><span class="sxs-lookup"><span data-stu-id="47169-136">Set the selected range</span></span>

<span data-ttu-id="47169-137">O exemplo de código a seguir seleciona o intervalo **B2:E6** na planilha ativa.</span><span class="sxs-lookup"><span data-stu-id="47169-137">The following code sample selects the range **B2:E6** in the active worksheet.</span></span>

```js
Excel.run(function (context) {
    var sheet = context.workbook.worksheets.getActiveWorksheet();
    var range = sheet.getRange("B2:E6");

    range.select();

    return context.sync();
}).catch(errorHandlerFunction);
```

<span data-ttu-id="47169-138">**Intervalo selecionado B2:E6**</span><span class="sxs-lookup"><span data-stu-id="47169-138">**Selected range B2:E6**</span></span>

![Intervalo selecionado no Excel](../images/excel-ranges-set-selection.png)

## <a name="get-the-selected-range"></a><span data-ttu-id="47169-140">Obter o intervalo selecionado</span><span class="sxs-lookup"><span data-stu-id="47169-140">Get the selected range</span></span>

<span data-ttu-id="47169-141">O exemplo de código a seguir obtém o intervalo selecionado, carrega sua propriedade **address** e grava uma mensagem no console.</span><span class="sxs-lookup"><span data-stu-id="47169-141">The following code sample gets the selected range, loads its **address** property, and writes a message to the console.</span></span> 

```js
Excel.run(function (context) {
    var range = context.workbook.getSelectedRange();
    range.load("address");

    return context.sync()
        .then(function () {
            console.log(`The address of the selected range is "${range.address}"`);
        });
}).catch(errorHandlerFunction);
```

## <a name="set-values-or-formulas"></a><span data-ttu-id="47169-142">Definir valores ou fórmulas</span><span class="sxs-lookup"><span data-stu-id="47169-142">Set values or formulas</span></span>

<span data-ttu-id="47169-143">Os exemplos a seguir mostram como definir valores e fórmulas para uma única célula ou um intervalo de células.</span><span class="sxs-lookup"><span data-stu-id="47169-143">The following examples show how to set values and formulas for a single cell or a range of cells.</span></span>

### <a name="set-value-for-a-single-cell"></a><span data-ttu-id="47169-144">Definir valor para uma única célula</span><span class="sxs-lookup"><span data-stu-id="47169-144">Set value for a single cell</span></span>

<span data-ttu-id="47169-145">O exemplo de código a seguir define o valor da célula **C3** como "5" e, em seguida, define a largura das colunas para melhor ajustar os dados.</span><span class="sxs-lookup"><span data-stu-id="47169-145">The following code sample sets the value of cell **C3** to "5" and then sets the width of the columns to best fit the data.</span></span>

```js
Excel.run(function (context) {
    var sheet = context.workbook.worksheets.getItem("Sample");

    var range = sheet.getRange("C3");
    range.values = [[ 5 ]];
    range.format.autofitColumns();

    return context.sync();
}).catch(errorHandlerFunction);
```

<span data-ttu-id="47169-146">**Dados antes da atualização do valor da célula**</span><span class="sxs-lookup"><span data-stu-id="47169-146">**Data before cell value is updated**</span></span>

![Dados no Excel antes da atualização do valor da célula](../images/excel-ranges-set-start.png)

<span data-ttu-id="47169-148">**Dados após a atualização do valor da célula**</span><span class="sxs-lookup"><span data-stu-id="47169-148">**Data after cell value is updated**</span></span>

![Dados no Excel após a atualização do valor da célula](../images/excel-ranges-set-cell-value.png)

### <a name="set-values-for-a-range-of-cells"></a><span data-ttu-id="47169-150">Definir valores para um intervalo de células</span><span class="sxs-lookup"><span data-stu-id="47169-150">Set values for a range of cells</span></span>

<span data-ttu-id="47169-151">O exemplo de código a seguir define valores das células no intervalo **B5:D5** e, em seguida, define a largura das colunas para melhor ajustar os dados.</span><span class="sxs-lookup"><span data-stu-id="47169-151">The following code sample sets values for the cells in the range **B5:D5** and then sets the width of the columns to best fit the data.</span></span>

```js
Excel.run(function (context) {
    var sheet = context.workbook.worksheets.getItem("Sample");

    var data = [
        ["Potato Chips", 10, 1.80],
    ];
    
    var range = sheet.getRange("B5:D5");
    range.values = data;
    range.format.autofitColumns();

    return context.sync();
}).catch(errorHandlerFunction);
```

<span data-ttu-id="47169-152">**Dados antes da atualização dos valores da célula**</span><span class="sxs-lookup"><span data-stu-id="47169-152">**Data before cell values are updated**</span></span>

![Dados no Excel antes da atualização dos valores da célula](../images/excel-ranges-set-start.png)

<span data-ttu-id="47169-154">**Dados após a atualização dos valores da célula**</span><span class="sxs-lookup"><span data-stu-id="47169-154">**Data after cell values are updated**</span></span>

![Dados no Excel após a atualização dos valores da célula](../images/excel-ranges-set-cell-values.png)

### <a name="set-formula-for-a-single-cell"></a><span data-ttu-id="47169-156">Definir fórmula para uma única célula</span><span class="sxs-lookup"><span data-stu-id="47169-156">Set formula for a single cell</span></span>

<span data-ttu-id="47169-157">O exemplo de código a seguir define uma fórmula para a célula **E3** e, em seguida, define a largura das colunas para melhor ajustar os dados.</span><span class="sxs-lookup"><span data-stu-id="47169-157">The following code sample sets a formula for cell **E3** and then sets the width of the columns to best fit the data.</span></span>

```js
Excel.run(function (context) {
    var sheet = context.workbook.worksheets.getItem("Sample");

    var range = sheet.getRange("E3");
    range.formulas = [[ "=C3 * D3" ]];
    range.format.autofitColumns();

    return context.sync();
}).catch(errorHandlerFunction);
```

<span data-ttu-id="47169-158">**Dados antes da definição da fórmula da célula**</span><span class="sxs-lookup"><span data-stu-id="47169-158">**Data before cell formula is set**</span></span>

![Dados no Excel antes da definição da fórmula da célula](../images/excel-ranges-start-set-formula.png)

<span data-ttu-id="47169-160">**Dados após a definição da fórmula da célula**</span><span class="sxs-lookup"><span data-stu-id="47169-160">**Data after cell formula is set**</span></span>

![Dados no Excel após a definição da fórmula da célula](../images/excel-ranges-set-formula.png)

### <a name="set-formulas-for-a-range-of-cells"></a><span data-ttu-id="47169-162">Definir fórmulas para um intervalo de células</span><span class="sxs-lookup"><span data-stu-id="47169-162">Set formulas for a range of cells</span></span>

<span data-ttu-id="47169-163">O exemplo de código a seguir define fórmulas para células no intervalo **E2:E6** e, em seguida, define a largura das colunas para melhor ajustar os dados.</span><span class="sxs-lookup"><span data-stu-id="47169-163">The following code sample sets formulas for cells in the range **E2:E6** and then sets the width of the columns to best fit the data.</span></span>

```js
Excel.run(function (context) {
    var sheet = context.workbook.worksheets.getItem("Sample");

    var data = [
        ["=C3 * D3"],
        ["=C4 * D4"],
        ["=C5 * D5"],
        ["=SUM(E3:E5)"]
    ];
    
    var range = sheet.getRange("E3:E6");
    range.formulas = data;
    range.format.autofitColumns();

    return context.sync();
}).catch(errorHandlerFunction);
```

<span data-ttu-id="47169-164">**Dados antes da definição das fórmulas da célula**</span><span class="sxs-lookup"><span data-stu-id="47169-164">**Data before cell formulas are set**</span></span>

![Dados no Excel antes da definição das fórmulas da célula](../images/excel-ranges-start-set-formula.png)

<span data-ttu-id="47169-166">**Dados após a definição das fórmulas da célula**</span><span class="sxs-lookup"><span data-stu-id="47169-166">**Data after cell formulas are set**</span></span>

![Dados no Excel após a definição das fórmulas da célula](../images/excel-ranges-set-formulas.png)

## <a name="get-values-text-or-formulas"></a><span data-ttu-id="47169-168">Obter valores, texto ou fórmulas</span><span class="sxs-lookup"><span data-stu-id="47169-168">Get values, text, or formulas</span></span>

<span data-ttu-id="47169-169">Estes exemplos mostram como obter valores, texto e fórmulas de um intervalo de células.</span><span class="sxs-lookup"><span data-stu-id="47169-169">These examples show how to get values, text, and formulas from a range of cells.</span></span>

### <a name="get-values-from-a-range-of-cells"></a><span data-ttu-id="47169-170">Obter valores de um intervalo de células</span><span class="sxs-lookup"><span data-stu-id="47169-170">Get values from a range of cells</span></span>

<span data-ttu-id="47169-171">O exemplo de código a seguir obtém o intervalo **B2:E6**, carrega sua propriedade **values** e grava os valores no console.</span><span class="sxs-lookup"><span data-stu-id="47169-171">The following code sample gets the range **B2:E6**, loads its **values** property, and writes the values to the console.</span></span> <span data-ttu-id="47169-172">A propriedade **values** de um intervalo especifica os novos valores brutos que as células contêm.</span><span class="sxs-lookup"><span data-stu-id="47169-172">The **values** property of a range specifies the raw values that the cells contain.</span></span> <span data-ttu-id="47169-173">Mesmo que algumas células em um intervalo contenham fórmulas, a propriedade **values** do intervalo especifica os valores brutos para essas células, não alguma das fórmulas.</span><span class="sxs-lookup"><span data-stu-id="47169-173">Even if some cells in a range contain formulas, the **values** property of the range specifies the raw values for those cells, not any of the formulas.</span></span>

```js
Excel.run(function (context) {
    var sheet = context.workbook.worksheets.getItem("Sample");
    var range = sheet.getRange("B2:E6");
    range.load("values");

    return context.sync()
        .then(function () {
            console.log(JSON.stringify(range.values, null, 4));
        });
}).catch(errorHandlerFunction);
```

<span data-ttu-id="47169-174">**Dados no intervalo (valores na coluna E são um resultado de fórmulas)**</span><span class="sxs-lookup"><span data-stu-id="47169-174">**Data in range (values in column E are a result of formulas)**</span></span>

![Dados no Excel após a definição das fórmulas da célula](../images/excel-ranges-set-formulas.png)

<span data-ttu-id="47169-176">**range.values (conforme registrado em log no console pelo exemplo de código acima)**</span><span class="sxs-lookup"><span data-stu-id="47169-176">**range.values (as logged to the console by the code sample above)**</span></span>

```json
[
    [
        "Product",
        "Qty",
        "Unit Price",
        "Total Price"
    ],
    [
        "Almonds",
        2,
        7.5,
        15
    ],
    [
        "Coffee",
        1,
        34.5,
        34.5
    ],
    [
        "Chocolate",
        5,
        9.56,
        47.8
    ],
    [
        "",
        "",
        "",
        97.3
    ]
]
```

### <a name="get-text-from-a-range-of-cells"></a><span data-ttu-id="47169-177">Obter texto de um intervalo de células</span><span class="sxs-lookup"><span data-stu-id="47169-177">Get text from a range of cells</span></span>

<span data-ttu-id="47169-178">O exemplo de código a seguir obtém o intervalo **B2:E6**, carrega sua propriedade **text** e o grava no console.</span><span class="sxs-lookup"><span data-stu-id="47169-178">The following code sample gets the range **B2:E6**, loads its **text** property, and writes it to the console.</span></span>  <span data-ttu-id="47169-179">A propriedade **text** de um intervalo especifica os valores de exibição para as células no intervalo.</span><span class="sxs-lookup"><span data-stu-id="47169-179">The **text** property of a range specifies the display values for cells in the range.</span></span> <span data-ttu-id="47169-180">Mesmo que algumas células em um intervalo contenham fórmulas, a propriedade **text** do intervalo especifica os valores de exibição para essas células, não alguma das fórmulas.</span><span class="sxs-lookup"><span data-stu-id="47169-180">Even if some cells in a range contain formulas, the **text** property of the range specifies the display values for those cells, not any of the formulas.</span></span>

```js
Excel.run(function (context) {
    var sheet = context.workbook.worksheets.getItem("Sample");
    var range = sheet.getRange("B2:E6");
    range.load("text");

    return context.sync()
        .then(function () {
            console.log(JSON.stringify(range.text, null, 4));
        });
}).catch(errorHandlerFunction);
```

<span data-ttu-id="47169-181">**Dados no intervalo (valores na coluna E são um resultado de fórmulas)**</span><span class="sxs-lookup"><span data-stu-id="47169-181">**Data in range (values in column E are a result of formulas)**</span></span>

![Dados no Excel após a definição das fórmulas da célula](../images/excel-ranges-set-formulas.png)

<span data-ttu-id="47169-183">**range.text (conforme registrado em log no console pelo exemplo de código acima)**</span><span class="sxs-lookup"><span data-stu-id="47169-183">**range.text (as logged to the console by the code sample above)**</span></span>

```json
[
    [
        "Product",
        "Qty",
        "Unit Price",
        "Total Price"
    ],
    [
        "Almonds",
        "2",
        "7.5",
        "15"
    ],
    [
        "Coffee",
        "1",
        "34.5",
        "34.5"
    ],
    [
        "Chocolate",
        "5",
        "9.56",
        "47.8"
    ],
    [
        "",
        "",
        "",
        "97.3"
    ]
]
```

### <a name="get-formulas-from-a-range-of-cells"></a><span data-ttu-id="47169-184">Obter fórmulas de um intervalo de células</span><span class="sxs-lookup"><span data-stu-id="47169-184">Get formulas from a range of cells</span></span>

<span data-ttu-id="47169-185">O exemplo de código a seguir obtém o intervalo **B2:E6**, carrega sua propriedade **formulas** e o grava no console.</span><span class="sxs-lookup"><span data-stu-id="47169-185">The following code sample gets the range **B2:E6**, loads its **formulas** property, and writes it to the console.</span></span>  <span data-ttu-id="47169-186">A propriedade **formulas** de um intervalo especifica as fórmulas para células no intervalo que contêm fórmulas e os valores brutos para células no intervalo que não contêm fórmulas.</span><span class="sxs-lookup"><span data-stu-id="47169-186">The **formulas** property of a range specifies the formulas for cells in the range that contain formulas and the raw values for cells in the range that do not contain formulas.</span></span>

```js
Excel.run(function (context) {
    var sheet = context.workbook.worksheets.getItem("Sample");
    var range = sheet.getRange("B2:E6");
    range.load("formulas");

    return context.sync()
        .then(function () {
            console.log(JSON.stringify(range.formulas, null, 4));
        });
}).catch(errorHandlerFunction);
```

<span data-ttu-id="47169-187">**Dados no intervalo (valores na coluna E são um resultado de fórmulas)**</span><span class="sxs-lookup"><span data-stu-id="47169-187">**Data in range (values in column E are a result of formulas)**</span></span>

![Dados no Excel após a definição das fórmulas da célula](../images/excel-ranges-set-formulas.png)

<span data-ttu-id="47169-189">**range.formulas (conforme registrado em log no console pelo exemplo de código acima)**</span><span class="sxs-lookup"><span data-stu-id="47169-189">**range.formulas (as logged to the console by the code sample above)**</span></span>

```json
[
    [
        "Product",
        "Qty",
        "Unit Price",
        "Total Price"
    ],
    [
        "Almonds",
        2,
        7.5,
        "=C3 * D3"
    ],
    [
        "Coffee",
        1,
        34.5,
        "=C4 * D4"
    ],
    [
        "Chocolate",
        5,
        9.56,
        "=C5 * D5"
    ],
    [
        "",
        "",
        "",
        "=SUM(E3:E5)"
    ]
]
```

## <a name="set-range-format"></a><span data-ttu-id="47169-190">Definir formato do intervalo</span><span class="sxs-lookup"><span data-stu-id="47169-190">Set range format</span></span>

<span data-ttu-id="47169-191">Os exemplos a seguir mostram como definir a cor da fonte, a cor de preenchimento e o formato de número para células em um intervalo.</span><span class="sxs-lookup"><span data-stu-id="47169-191">The following examples show how to set font color, fill color, and number format for cells in a range.</span></span>

### <a name="set-font-color-and-fill-color"></a><span data-ttu-id="47169-192">Definir cor da fonte e cor de preenchimento</span><span class="sxs-lookup"><span data-stu-id="47169-192">Set font color and fill color</span></span>

<span data-ttu-id="47169-193">O exemplo de código a seguir define a cor da fonte e a cor de preenchimento para células no intervalo **B2:E2**.</span><span class="sxs-lookup"><span data-stu-id="47169-193">The following code sample sets the font color and fill color for cells in range **B2:E2**.</span></span>

```js
Excel.run(function (context) {
    var sheet = context.workbook.worksheets.getItem("Sample");

    var range = sheet.getRange("B2:E2");
    range.format.fill.color = "#4472C4";;
    range.format.font.color = "white";

    return context.sync();
}).catch(errorHandlerFunction);
```

<span data-ttu-id="47169-194">**Dados no intervalo antes da definição da cor da fonte e da cor de preenchimento**</span><span class="sxs-lookup"><span data-stu-id="47169-194">**Data in range before font color and fill color are set**</span></span>

![Dados no Excel antes da definição do formato](../images/excel-ranges-format-before.png)

<span data-ttu-id="47169-196">**Dados no intervalo após a definição da cor da fonte e da cor de preenchimento**</span><span class="sxs-lookup"><span data-stu-id="47169-196">**Data in range after font color and fill color are set**</span></span>

![Dados no Excel após a definição do formato](../images/excel-ranges-format-font-and-fill.png)

### <a name="set-number-format"></a><span data-ttu-id="47169-198">Definir formato de número</span><span class="sxs-lookup"><span data-stu-id="47169-198">Set number format</span></span>

<span data-ttu-id="47169-199">O exemplo de código a seguir define o formato de número para as células no intervalo **D3:E5**.</span><span class="sxs-lookup"><span data-stu-id="47169-199">The following code sample sets the number format for the cells in range **D3:E5**.</span></span>

```js
Excel.run(function (context) {
    var sheet = context.workbook.worksheets.getItem("Sample");

    var formats = [
        ["0.00", "0.00"],
        ["0.00", "0.00"],
        ["0.00", "0.00"]
    ];

    var range = sheet.getRange("D3:E5");
    range.numberFormat = formats;

    return context.sync();
}).catch(errorHandlerFunction);
```

<span data-ttu-id="47169-200">**Dados no intervalo antes da definição do formato de número**</span><span class="sxs-lookup"><span data-stu-id="47169-200">**Data in range before number format is set**</span></span>

![Dados no Excel antes da definição do formato](../images/excel-ranges-format-font-and-fill.png)

<span data-ttu-id="47169-202">**Dados no intervalo após a definição do formato de número**</span><span class="sxs-lookup"><span data-stu-id="47169-202">**Data in range after number format is set**</span></span>

![Dados no Excel após a definição do formato](../images/excel-ranges-format-numbers.png)

### <a name="conditional-formatting-of-ranges"></a><span data-ttu-id="47169-204">Formatação condicional de intervalos</span><span class="sxs-lookup"><span data-stu-id="47169-204">Conditional formatting of ranges</span></span>

<span data-ttu-id="47169-205">Os intervalos podem ter formatos aplicados a células individuais baseadas em condições.</span><span class="sxs-lookup"><span data-stu-id="47169-205">Ranges can have formats applied to individual cells based on conditions.</span></span> <span data-ttu-id="47169-206">Confira mais informações sobre isso em [Aplicar a formatação condicional a intervalos do Excel](excel-add-ins-conditional-formatting.md).</span><span class="sxs-lookup"><span data-stu-id="47169-206">For more information about this, see [Apply conditional formatting to Excel ranges](excel-add-ins-conditional-formatting.md).</span></span>

## <a name="find-a-cell-using-string-matching-preview"></a><span data-ttu-id="47169-207">Localizar uma célula usando a cadeia de correspondência (versão prévia)</span><span class="sxs-lookup"><span data-stu-id="47169-207">Find a cell using string matching (preview)</span></span>

> [!NOTE]
> <span data-ttu-id="47169-208">A função `find` do objeto do intervalo só está disponível atualmente na versão prévia pública.</span><span class="sxs-lookup"><span data-stu-id="47169-208">The Range object's `find` function is currently available only in public preview (beta).</span></span> [!INCLUDE [Information about using preview APIs](../includes/using-preview-apis.md)]

<span data-ttu-id="47169-209">O objeto `Range` tem um método `find` para pesquisar uma cadeia especificada dentro do intervalo.</span><span class="sxs-lookup"><span data-stu-id="47169-209">The `Range` object has a `find` method to search for a specified string within the range.</span></span> <span data-ttu-id="47169-210">Ele retorna o intervalo da primeira célula com o texto correspondente.</span><span class="sxs-lookup"><span data-stu-id="47169-210">It returns the range of the first cell with matching text.</span></span> <span data-ttu-id="47169-211">O exemplo de código a seguir localiza a primeira célula com um valor igual à cadeia de caracteres **Alimentos** e registra o seu endereço no console.</span><span class="sxs-lookup"><span data-stu-id="47169-211">The following code sample finds the first cell with a value equal to the string **Food** and logs its address to the console.</span></span> <span data-ttu-id="47169-212">Observe que `find` exibe um erro `ItemNotFound` se a cadeia de caracteres especificada não existir no intervalo.</span><span class="sxs-lookup"><span data-stu-id="47169-212">Note that `find` throws an `ItemNotFound` error if the specified string doesn't exist in the range.</span></span> <span data-ttu-id="47169-213">Se você acha que a cadeia de caracteres especificada pode não estar no intervalo, use o método [findOrNullObject](excel-add-ins-advanced-concepts.md#42ornullobject-methods) para que seu código manipule normalmente esse cenário.</span><span class="sxs-lookup"><span data-stu-id="47169-213">If you expect that the specified string may not exist in the range, use the [findOrNullObject](excel-add-ins-advanced-concepts.md#42ornullobject-methods) method instead, so your code gracefully handles that scenario.</span></span>

```js
Excel.run(function (context) {
    var sheet = context.workbook.worksheets.getItem("Sample");
    var table = sheet.tables.getItem("ExpensesTable");
    var searchRange = table.getRange();
    var foundRange = searchRange.find("Food", {
        completeMatch: true, // find will match the whole cell value
        matchCase: false, // find will not match case
        searchDirection: Excel.SearchDirection.forward // find will start searching at the beginning of the range
    });

    foundRange.load("address");
    return context.sync()
        .then(function() {
            console.log(foundRange.address);
    });
}).catch(errorHandlerFunction);
```

<span data-ttu-id="47169-214">Quando o método `find` é chamado em um intervalo que representa uma única célula, a planilha inteira é pesquisada.</span><span class="sxs-lookup"><span data-stu-id="47169-214">When the `find` method is called on a range representing a single cell, the entire worksheet is searched.</span></span> <span data-ttu-id="47169-215">A pesquisa começa na célula e segue na direção especificada pelo `SearchCriteria.searchDirection`, envolvendo as extremidades da planilha, se necessário.</span><span class="sxs-lookup"><span data-stu-id="47169-215">The search begins at that cell and goes in the direction specified by `SearchCriteria.searchDirection`, wrapping around the ends of the worksheet if needed.</span></span>

## <a name="see-also"></a><span data-ttu-id="47169-216">Confira também</span><span class="sxs-lookup"><span data-stu-id="47169-216">See also</span></span>

- [<span data-ttu-id="47169-217">Trabalhar com intervalos usando a API JavaScript do Excel (avançado)</span><span class="sxs-lookup"><span data-stu-id="47169-217">Work with ranges using the Excel JavaScript API (advanced)</span></span>](excel-add-ins-ranges-advanced.md)
- [<span data-ttu-id="47169-218">Conceitos fundamentais de programação com a API JavaScript do Excel</span><span class="sxs-lookup"><span data-stu-id="47169-218">Fundamental programming concepts with the Excel JavaScript API</span></span>](excel-add-ins-core-concepts.md)
