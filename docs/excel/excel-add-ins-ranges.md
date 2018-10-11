---
title: Trabalhar com intervalos usando a API JavaScript do Excel
description: ''
ms.date: 12/04/2017
ms.openlocfilehash: 246b882a921b5a43ca747238262af7c4b23c97ee
ms.sourcegitcommit: 563c53bac52b31277ab935f30af648f17c5ed1e2
ms.translationtype: HT
ms.contentlocale: pt-BR
ms.lasthandoff: 10/10/2018
ms.locfileid: "25459165"
---
# <a name="work-with-ranges-using-the-excel-javascript-api"></a><span data-ttu-id="22038-102">Trabalhar com intervalos usando a API JavaScript do Excel</span><span class="sxs-lookup"><span data-stu-id="22038-102">Work with ranges using the Excel JavaScript API</span></span>

<span data-ttu-id="22038-p101">Este artigo fornece exemplos de código que mostram como realizar tarefas comuns com intervalos usando a API JavaScript do Excel. Para obter uma lista completa de propriedades e métodos que o objeto **Range**  suporta, confira [Objeto Range (JavaScript API para Excel)](https://docs.microsoft.com/javascript/api/excel/excel.range).</span><span class="sxs-lookup"><span data-stu-id="22038-p101">This article provides code samples that show how to perform common tasks with ranges using the Excel JavaScript API. For the complete list of properties and methods that the **Range** object supports, see [Range Object (JavaScript API for Excel)](https://docs.microsoft.com/javascript/api/excel/excel.range).</span></span>

## <a name="get-a-range"></a><span data-ttu-id="22038-105">Obter um intervalo</span><span class="sxs-lookup"><span data-stu-id="22038-105">Get a range</span></span>

<span data-ttu-id="22038-106">Os exemplos a seguir mostram diferentes maneiras de obter uma referência a um intervalo em uma planilha.</span><span class="sxs-lookup"><span data-stu-id="22038-106">The following examples show different ways to get a reference to a range within a worksheet.</span></span>

### <a name="get-range-by-address"></a><span data-ttu-id="22038-107">Obter intervalo por endereço</span><span class="sxs-lookup"><span data-stu-id="22038-107">Get range by address</span></span>

<span data-ttu-id="22038-108">O exemplo de código a seguir obtém o intervalo com o endereço **B2:B5** da planilha chamada **Sample**, carrega sua propriedade **address** e grava uma mensagem no console.</span><span class="sxs-lookup"><span data-stu-id="22038-108">The following code sample gets the range with address **B2:B5** from the worksheet named **Sample**, loads its **address** property, and writes a message to the console.</span></span>

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

### <a name="get-range-by-name"></a><span data-ttu-id="22038-109">Obter intervalo por nome</span><span class="sxs-lookup"><span data-stu-id="22038-109">Get range by name</span></span>

<span data-ttu-id="22038-110">O exemplo de código a seguir obtém o intervalo chamado **MyRange** da planilha chamada **Sample**, carrega sua propriedade **address** e grava uma mensagem no console.</span><span class="sxs-lookup"><span data-stu-id="22038-110">The following code sample gets the range named **MyRange** from the worksheet named **Sample**, loads its **address** property, and writes a message to the console.</span></span>

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

### <a name="get-used-range"></a><span data-ttu-id="22038-111">Obter intervalo usado</span><span class="sxs-lookup"><span data-stu-id="22038-111">Get used range</span></span>

<span data-ttu-id="22038-p102">O exemplo de código a seguir obtém o intervalo usado da planilha chamada **Sample**, carrega sua propriedade de **address** e grava uma mensagem no console. O intervalo usado é o menor intervalo que abrange quaisquer células na planilha que tenham um valor ou formatação atribuída a elas. Se a planilha inteira estiver em branco, o método **getUsedRange()** retornará um intervalo que consiste apenas na célula superior esquerda da planilha.</span><span class="sxs-lookup"><span data-stu-id="22038-p102">The following code sample gets the used range from the worksheet named **Sample**, loads its **address** property, and writes a message to the console. The used range is the smallest range that encompasses any cells in the worksheet that have a value or formatting assigned to them. If the entire worksheet is blank, the **getUsedRange()** method returns a range that consists of only the top-left cell in the worksheet.</span></span>

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

### <a name="get-entire-range"></a><span data-ttu-id="22038-115">Obter intervalo inteiro</span><span class="sxs-lookup"><span data-stu-id="22038-115">Get entire range</span></span>

<span data-ttu-id="22038-116">O exemplo de código a seguir obtém todo o intervalo da planilha chamada **Sample**, carrega sua propriedade **address** e grava uma mensagem no console.</span><span class="sxs-lookup"><span data-stu-id="22038-116">The following code sample gets the entire worksheet range from the worksheet named **Sample**, loads its **address** property, and writes a message to the console.</span></span>

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

## <a name="insert-a-range-of-cells"></a><span data-ttu-id="22038-117">Inserir um intervalo de células</span><span class="sxs-lookup"><span data-stu-id="22038-117">Insert a range of cells</span></span>

<span data-ttu-id="22038-118">O exemplo de código a seguir insere um intervalo de células no local **B4:E4** e desloca outras células para baixo a fim de fornecer espaço para as novas células.</span><span class="sxs-lookup"><span data-stu-id="22038-118">The following code sample inserts a range of cells in location **B4:E4** and shifts other cells down to provide space for the new cells.</span></span>

```js
Excel.run(function (context) {
    var sheet = context.workbook.worksheets.getItem("Sample");
    var range = sheet.getRange("B4:E4");

    range.insert(Excel.InsertShiftDirection.down);
    
    return context.sync();
}).catch(errorHandlerFunction);
```

<span data-ttu-id="22038-119">**Dados antes da inserção do intervalo**</span><span class="sxs-lookup"><span data-stu-id="22038-119">**Data before range is inserted**</span></span>

![Dados no Excel antes da inserção do intervalo](../images/excel-ranges-start.png)

<span data-ttu-id="22038-121">**Dados após a inserção do intervalo**</span><span class="sxs-lookup"><span data-stu-id="22038-121">**Data after range is inserted**</span></span>

![Dados no Excel após a inserção do intervalo](../images/excel-ranges-after-insert.png)

## <a name="clear-a-range-of-cells"></a><span data-ttu-id="22038-123">Limpar um intervalo de células</span><span class="sxs-lookup"><span data-stu-id="22038-123">Clear a range of cells</span></span>

<span data-ttu-id="22038-124">O exemplo de código a seguir limpa todo o conteúdo e a formatação das células no intervalo **E2:E5**.</span><span class="sxs-lookup"><span data-stu-id="22038-124">The following code sample clears all contents and formatting of cells in the range **E2:E5**.</span></span>  

```js
Excel.run(function (context) {
    var sheet = context.workbook.worksheets.getItem("Sample");
    var range = sheet.getRange("E2:E5");

    range.clear();

    return context.sync();
}).catch(errorHandlerFunction);
```

<span data-ttu-id="22038-125">**Dados antes da limpeza do intervalo**</span><span class="sxs-lookup"><span data-stu-id="22038-125">**Data before range is cleared**</span></span>

![Dados no Excel antes da limpeza do intervalo](../images/excel-ranges-start.png)

<span data-ttu-id="22038-127">**Dados após a limpeza do intervalo**</span><span class="sxs-lookup"><span data-stu-id="22038-127">**Data after range is cleared**</span></span>

![Dados no Excel após a limpeza do intervalo](../images/excel-ranges-after-clear.png)

## <a name="delete-a-range-of-cells"></a><span data-ttu-id="22038-129">Excluir um intervalo de células</span><span class="sxs-lookup"><span data-stu-id="22038-129">Delete a range of cells</span></span>

<span data-ttu-id="22038-130">O exemplo de código a seguir exclui as células no intervalo **B4:E4** e desloca outras células para cima a fim de preencher o espaço deixado pelas células excluídas.</span><span class="sxs-lookup"><span data-stu-id="22038-130">The following code sample deletes the cells in the range **B4:E4** and shift other cells up to fill the space that was vacated by the deleted cells.</span></span>

```js
Excel.run(function (context) {
    var sheet = context.workbook.worksheets.getItem("Sample");
    var range = sheet.getRange("B4:E4");

    range.delete(Excel.DeleteShiftDirection.up);

    return context.sync();
}).catch(errorHandlerFunction);
```

<span data-ttu-id="22038-131">**Dados antes da exclusão do intervalo**</span><span class="sxs-lookup"><span data-stu-id="22038-131">**Data before range is deleted**</span></span>

![Dados no Excel antes da exclusão do intervalo](../images/excel-ranges-start.png)

<span data-ttu-id="22038-133">**Dados após a exclusão do intervalo**</span><span class="sxs-lookup"><span data-stu-id="22038-133">**Data after range is deleted**</span></span>

![Dados no Excel após a exclusão do intervalo](../images/excel-ranges-after-delete.png)

## <a name="set-the-selected-range"></a><span data-ttu-id="22038-135">Definir o intervalo selecionado</span><span class="sxs-lookup"><span data-stu-id="22038-135">Set the selected range</span></span>

<span data-ttu-id="22038-136">O exemplo de código a seguir seleciona o intervalo **B2:E6** na planilha ativa.</span><span class="sxs-lookup"><span data-stu-id="22038-136">The following code sample selects the range **B2:E6** in the active worksheet.</span></span>

```js
Excel.run(function (context) {
    var sheet = context.workbook.worksheets.getActiveWorksheet();
    var range = sheet.getRange("B2:E6");

    range.select();

    return context.sync();
}).catch(errorHandlerFunction);
```

<span data-ttu-id="22038-137">**Intervalo selecionado B2:E6**</span><span class="sxs-lookup"><span data-stu-id="22038-137">**Selected range B2:E6**</span></span>

![Intervalo selecionado no Excel](../images/excel-ranges-set-selection.png)

## <a name="get-the-selected-range"></a><span data-ttu-id="22038-139">Obter o intervalo selecionado</span><span class="sxs-lookup"><span data-stu-id="22038-139">Get the selected range</span></span>

<span data-ttu-id="22038-140">O exemplo de código a seguir obtém o intervalo selecionado, carrega sua propriedade **address** e grava uma mensagem no console.</span><span class="sxs-lookup"><span data-stu-id="22038-140">The following code sample gets the selected range, loads its **address** property, and writes a message to the console.</span></span> 

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

## <a name="set-values-or-formulas"></a><span data-ttu-id="22038-141">Definir valores ou fórmulas</span><span class="sxs-lookup"><span data-stu-id="22038-141">Set values or formulas</span></span>

<span data-ttu-id="22038-142">Os exemplos a seguir mostram como atrubuir valores e fórmulas para uma única célula ou um intervalo de células.</span><span class="sxs-lookup"><span data-stu-id="22038-142">The following examples show how to set values and formulas for a single cell or a range of cells.</span></span>

### <a name="set-value-for-a-single-cell"></a><span data-ttu-id="22038-143">Definir valor para uma única célula</span><span class="sxs-lookup"><span data-stu-id="22038-143">Set value for a single cell</span></span>

<span data-ttu-id="22038-144">O exemplo de código a seguir atribui o valor da célula **C3** como "5" e, em seguida, define a largura das colunas para melhor ajustar os dados.</span><span class="sxs-lookup"><span data-stu-id="22038-144">The following code sample sets the value of cell **C3** to "5" and then sets the width of the columns to best fit the data.</span></span>

```js
Excel.run(function (context) {
    var sheet = context.workbook.worksheets.getItem("Sample");

    var range = sheet.getRange("C3");
    range.values = [[ 5 ]];
    range.format.autofitColumns();

    return context.sync();
}).catch(errorHandlerFunction);
```

<span data-ttu-id="22038-145">**Dados antes da atualização do valor da célula**</span><span class="sxs-lookup"><span data-stu-id="22038-145">**Data before cell value is updated**</span></span>

![Dados no Excel antes da atualização do valor da célula](../images/excel-ranges-set-start.png)

<span data-ttu-id="22038-147">**Dados após a atualização do valor da célula**</span><span class="sxs-lookup"><span data-stu-id="22038-147">**Data after cell value is updated**</span></span>

![Dados no Excel após a atualização do valor da célula](../images/excel-ranges-set-cell-value.png)

### <a name="set-values-for-a-range-of-cells"></a><span data-ttu-id="22038-149">Definir valores para um intervalo de células</span><span class="sxs-lookup"><span data-stu-id="22038-149">Set values for a range of cells</span></span>

<span data-ttu-id="22038-150">O exemplo de código a seguir define valores das células no intervalo **B5:D5** e, em seguida, define a largura das colunas para melhor ajustar os dados.</span><span class="sxs-lookup"><span data-stu-id="22038-150">The following code sample sets values for the cells in the range **B5:D5** and then sets the width of the columns to best fit the data.</span></span>

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

<span data-ttu-id="22038-151">**Dados antes da atualização dos valores da célula**</span><span class="sxs-lookup"><span data-stu-id="22038-151">**Data before cell values are updated**</span></span>

![Dados no Excel antes da atualização dos valores da célula](../images/excel-ranges-set-start.png)

<span data-ttu-id="22038-153">**Dados após a atualização dos valores da célula**</span><span class="sxs-lookup"><span data-stu-id="22038-153">**Data after cell values are updated**</span></span>

![Dados no Excel após a atualização dos valores da célula](../images/excel-ranges-set-cell-values.png)

### <a name="set-formula-for-a-single-cell"></a><span data-ttu-id="22038-155">Definir fórmula para uma única célula</span><span class="sxs-lookup"><span data-stu-id="22038-155">Set formula for a single cell</span></span>

<span data-ttu-id="22038-156">O exemplo de código a seguir define uma fórmula para a célula **E3** e, em seguida, define a largura das colunas para melhor ajustar os dados.</span><span class="sxs-lookup"><span data-stu-id="22038-156">The following code sample sets a formula for cell **E3** and then sets the width of the columns to best fit the data.</span></span>

```js
Excel.run(function (context) {
    var sheet = context.workbook.worksheets.getItem("Sample");

    var range = sheet.getRange("E3");
    range.formulas = [[ "=C3 * D3" ]];
    range.format.autofitColumns();

    return context.sync();
}).catch(errorHandlerFunction);
```

<span data-ttu-id="22038-157">**Dados antes da definição da fórmula da célula**</span><span class="sxs-lookup"><span data-stu-id="22038-157">**Data before cell formula is set**</span></span>

![Dados no Excel antes da definição da fórmula da célula](../images/excel-ranges-start-set-formula.png)

<span data-ttu-id="22038-159">**Dados após a definição da fórmula da célula**</span><span class="sxs-lookup"><span data-stu-id="22038-159">**Data after cell formula is set**</span></span>

![Dados no Excel após a definição da fórmula da célula](../images/excel-ranges-set-formula.png)

### <a name="set-formulas-for-a-range-of-cells"></a><span data-ttu-id="22038-161">Definir fórmulas para um intervalo de células</span><span class="sxs-lookup"><span data-stu-id="22038-161">Set formulas for a range of cells</span></span>

<span data-ttu-id="22038-162">O exemplo de código a seguir define fórmulas para células no intervalo **E2:E6** e, em seguida, define a largura das colunas para melhor ajustar os dados.</span><span class="sxs-lookup"><span data-stu-id="22038-162">The following code sample sets formulas for cells in the range **E2:E6** and then sets the width of the columns to best fit the data.</span></span>

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

<span data-ttu-id="22038-163">**Dados antes da definição das fórmulas da célula**</span><span class="sxs-lookup"><span data-stu-id="22038-163">**Data before cell formulas are set**</span></span>

![Dados no Excel antes da definição das fórmulas da célula](../images/excel-ranges-start-set-formula.png)

<span data-ttu-id="22038-165">**Dados após a definição das fórmulas da célula**</span><span class="sxs-lookup"><span data-stu-id="22038-165">**Data after cell formulas are set**</span></span>

![Dados no Excel após a definição das fórmulas da célula](../images/excel-ranges-set-formulas.png)

## <a name="get-values-text-or-formulas"></a><span data-ttu-id="22038-167">Obter valores, texto ou fórmulas</span><span class="sxs-lookup"><span data-stu-id="22038-167">Get values, text, or formulas</span></span>

<span data-ttu-id="22038-168">Estes exemplos mostram como obter valores, texto e fórmulas de um intervalo de células.</span><span class="sxs-lookup"><span data-stu-id="22038-168">These examples show how to get values, text, and formulas from a range of cells.</span></span>

### <a name="get-values-from-a-range-of-cells"></a><span data-ttu-id="22038-169">Obter valores de um intervalo de células</span><span class="sxs-lookup"><span data-stu-id="22038-169">Get values from a range of cells</span></span>

<span data-ttu-id="22038-p103">O exemplo de código a seguir obtém o intervalo **B2:E6**, carrega sua propriedade **values** e grava os valores no console. A propriedade **values** de um intervalo especifica os valores brutos que as células contêm. Mesmo que algumas células em um intervalo contenham fórmulas, a propriedade **values** do intervalo especifica os valores brutos para essas células, não para nenhuma das fórmulas.</span><span class="sxs-lookup"><span data-stu-id="22038-p103">The following code sample gets the range **B2:E6**, loads its **values** property, and writes the values to the console. The **values** property of a range specifies the raw values that the cells contain. Even if some cells in a range contain formulas, the **values** property of the range specifies the raw values for those cells, not any of the formulas.</span></span>

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

<span data-ttu-id="22038-173">**Dados no intervalo (valores na coluna E são um resultado de fórmulas)**</span><span class="sxs-lookup"><span data-stu-id="22038-173">**Data in range (values in column E are a result of formulas)**</span></span>

![Dados no Excel após a definição das fórmulas da célula](../images/excel-ranges-set-formulas.png)

<span data-ttu-id="22038-175">**range.values (conforme registrado em log no console pelo exemplo de código acima)**</span><span class="sxs-lookup"><span data-stu-id="22038-175">**range.values (as logged to the console by the code sample above)**</span></span>

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

### <a name="get-text-from-a-range-of-cells"></a><span data-ttu-id="22038-176">Obter texto de um intervalo de células</span><span class="sxs-lookup"><span data-stu-id="22038-176">Get text from a range of cells</span></span>

<span data-ttu-id="22038-p104">O exemplo de código a seguir obtém o intervalo **B2:E6**, carrega sua propriedade **text** e a grava no console. A propriedade **text** de um intervalo especifica os valores de exibição para células no intervalo. Mesmo que algumas células em um intervalo contenham fórmulas, a propriedade **text** do intervalo especifica os valores de exibição para essas células, não qualquer uma das fórmulas.</span><span class="sxs-lookup"><span data-stu-id="22038-p104">The following code sample gets the range **B2:E6**, loads its **text** property, and writes it to the console.  The **text** property of a range specifies the display values for cells in the range. Even if some cells in a range contain formulas, the **text** property of the range specifies the display values for those cells, not any of the formulas.</span></span>

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

<span data-ttu-id="22038-180">**Dados no intervalo (valores na coluna E são um resultado de fórmulas)**</span><span class="sxs-lookup"><span data-stu-id="22038-180">**Data in range (values in column E are a result of formulas)**</span></span>

![Dados no Excel após a definição das fórmulas da célula](../images/excel-ranges-set-formulas.png)

<span data-ttu-id="22038-182">**range.text (conforme registrado em log no console pelo exemplo de código acima)**</span><span class="sxs-lookup"><span data-stu-id="22038-182">**range.text (as logged to the console by the code sample above)**</span></span>

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

### <a name="get-formulas-from-a-range-of-cells"></a><span data-ttu-id="22038-183">Obter fórmulas de um intervalo de células</span><span class="sxs-lookup"><span data-stu-id="22038-183">Get formulas from a range of cells</span></span>

<span data-ttu-id="22038-p105">O exemplo de código a seguir obtém o intervalo **B2:E6**, carrega sua propriedade **formulas** e a grava no console.  A propriedade **formulas** de um intervalo especifica as fórmulas para células no intervalo que contêm fórmulas e os valores brutos para células no intervalo que não contêm fórmulas.</span><span class="sxs-lookup"><span data-stu-id="22038-p105">The following code sample gets the range **B2:E6**, loads its **formulas** property, and writes it to the console.  The **formulas** property of a range specifies the formulas for cells in the range that contain formulas and the raw values for cells in the range that do not contain formulas.</span></span>

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

<span data-ttu-id="22038-186">**Dados no intervalo (valores na coluna E são um resultado de fórmulas)**</span><span class="sxs-lookup"><span data-stu-id="22038-186">**Data in range (values in column E are a result of formulas)**</span></span>

![Dados no Excel após a definição das fórmulas da célula](../images/excel-ranges-set-formulas.png)

<span data-ttu-id="22038-188">**range.formulas (conforme registrado em log no console pelo exemplo de código acima)**</span><span class="sxs-lookup"><span data-stu-id="22038-188">**range.formulas (as logged to the console by the code sample above)**</span></span>

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

## <a name="set-range-format"></a><span data-ttu-id="22038-189">Definir formato do intervalo</span><span class="sxs-lookup"><span data-stu-id="22038-189">Set range format</span></span>

<span data-ttu-id="22038-190">Os exemplos a seguir mostram como definir cor de fonte, cor de preenchimento e formato de número para células em um intervalo.</span><span class="sxs-lookup"><span data-stu-id="22038-190">The following examples show how to set font color, fill color, and number format for cells in a range.</span></span>

### <a name="set-font-color-and-fill-color"></a><span data-ttu-id="22038-191">Definir cor da fonte e cor de preenchimento</span><span class="sxs-lookup"><span data-stu-id="22038-191">Set font color and fill color</span></span>

<span data-ttu-id="22038-192">O exemplo de código a seguir define a cor da fonte e a cor de preenchimento das células no intervalo **B2:E2**.</span><span class="sxs-lookup"><span data-stu-id="22038-192">The following code sample sets the font color and fill color for cells in range **B2:E2**.</span></span>

```js
Excel.run(function (context) {
    var sheet = context.workbook.worksheets.getItem("Sample");

    var range = sheet.getRange("B2:E2");
    range.format.fill.color = "#4472C4";;
    range.format.font.color = "white";

    return context.sync();
}).catch(errorHandlerFunction);
```

<span data-ttu-id="22038-193">**Dados no intervalo antes da definição da cor da fonte e da cor de preenchimento**</span><span class="sxs-lookup"><span data-stu-id="22038-193">**Data in range before font color and fill color are set**</span></span>

![Dados no Excel antes da definição do formato](../images/excel-ranges-format-before.png)

<span data-ttu-id="22038-195">**Dados no intervalo após a definição da cor da fonte e da cor de preenchimento**</span><span class="sxs-lookup"><span data-stu-id="22038-195">**Data in range after font color and fill color are set**</span></span>

![Dados no Excel após a definição do formato](../images/excel-ranges-format-font-and-fill.png)

### <a name="set-number-format"></a><span data-ttu-id="22038-197">Definir formato de número</span><span class="sxs-lookup"><span data-stu-id="22038-197">Set number format</span></span>

<span data-ttu-id="22038-198">O exemplo de código a seguir define o formato de número para as células no intervalo **D3:E5**.</span><span class="sxs-lookup"><span data-stu-id="22038-198">The following code sample sets the number format for the cells in range **D3:E5**.</span></span>

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

<span data-ttu-id="22038-199">**Dados no intervalo antes da definição do formato de número**</span><span class="sxs-lookup"><span data-stu-id="22038-199">**Data in range before number format is set**</span></span>

![Dados no Excel antes da definição do formato](../images/excel-ranges-format-font-and-fill.png)

<span data-ttu-id="22038-201">**Dados no intervalo após a definição do formato de número**</span><span class="sxs-lookup"><span data-stu-id="22038-201">**Data in range after number format is set**</span></span>

![Dados no Excel após a definição do formato](../images/excel-ranges-format-numbers.png)

## <a name="copy-and-paste"></a><span data-ttu-id="22038-203">Copiar e colar</span><span class="sxs-lookup"><span data-stu-id="22038-203">Copy and paste</span></span>

> [!NOTE]
> <span data-ttu-id="22038-p106">A função copyFrom está atualmente disponível somente na visualização pública (beta). Para usar esse recurso, você deve usar a biblioteca de beta do CDN Office. js: https://appsforoffice.microsoft.com/lib/beta/hosted/office.js. Se você estiver usando o TypeScript ou o seu editor de código usar arquivos de definição de tipo TypeScript para o IntelliSense, use https://appsforoffice.microsoft.com/lib/beta/hosted/office.d.ts.</span><span class="sxs-lookup"><span data-stu-id="22038-p106">The copyFrom function is currently available only in public preview (beta). To use this feature, you must use the beta library of the Office.js CDN: https://appsforoffice.microsoft.com/lib/beta/hosted/office.js. If you are using TypeScript or your code editor uses TypeScript type definition files for IntelliSense, use https://appsforoffice.microsoft.com/lib/beta/hosted/office.d.ts.</span></span>

<span data-ttu-id="22038-p107">A função copyFrom do intervalo replica o comportamento de copiar e colar da interface do Excel. O objeto de intervalo em que o copyFrom é chamado é o destino. A origem a ser copiada é passada como um intervalo ou um endereço de cadeia representando um intervalo. O exemplo de código a seguir copia os dados de **a1: E1** para o intervalo começando em **G1** (que acaba colando no **G1:K1**).</span><span class="sxs-lookup"><span data-stu-id="22038-p107">Range’s copyFrom function replicates the copy-and-paste behavior of the Excel UI. The range object that copyFrom is called on is the destination. The source to be copied is passed as a range or a string address representing a range. The following code sample copies the data from **A1:E1** into the range starting at **G1** (which ends up pasting into **G1:K1**).</span></span>

```js
Excel.run(function (context) {
    var sheet = context.workbook.worksheets.getItem("Sample");
    // copy a range starting at a single cell destination
    sheet.getRange("G1").copyFrom("A1:E1");
    return context.sync();
}).catch(errorHandlerFunction);
```

<span data-ttu-id="22038-211">Range.copyFrom tem três parâmetros opcionais.</span><span class="sxs-lookup"><span data-stu-id="22038-211">Range.copyFrom has three optional parameters.</span></span>

```ts
copyFrom(sourceRange: Range | string, copyType?: "All" | "Formulas" | "Values" | "Formats", skipBlanks?: boolean, transpose?: boolean): void;
``` 

<span data-ttu-id="22038-p108">`copyType` especifica quais dados são copiados da origem para o destino. `“Formulas”` transfere as fórmulas nas células de origem e preserva o posicionamento relativo dos intervalos dessas fórmulas. Quaisquer entradas que não sejam fórmulas são copiadas como estão. `“Values”` copia os valores de dados e, no caso de fórmulas, o resultado da fórmula. `“Formats”` copia a formatação do intervalo, incluindo fonte, cor e outras configurações de formato, mas sem valores. `”All”` (a opção padrão) copia os dados e a formatação, preservando as fórmulas das células, se encontradas.</span><span class="sxs-lookup"><span data-stu-id="22038-p108">`copyType` specifies what data gets copied from the source to the destination. `“Formulas”` transfers the formulas in the source cells and preserves the relative positioning of those formulas’ ranges. Any non-formula entries are copied as-is. `“Values”` copies the data values and, in the case of formulas, the result of the formula. `“Formats”` copies the formatting of the range, including font, color, and other format settings, but no values. `”All”` (the default option) copies both data and formatting, preserving cells’ formulas if found.</span></span>

<span data-ttu-id="22038-p109">`skipBlanks` define se as células em branco são copiadas para o destino. Quando verdadeiro, `copyFrom` ignora as células em branco no intervalo de origem. As células ignoradas não sobrescreverão os dados existentes de suas células correspondentes no intervalo de destino. O padrão é falso.</span><span class="sxs-lookup"><span data-stu-id="22038-p109">`skipBlanks` sets whether blank cells are copied into the destination. When true, `copyFrom` skips blank cells in the source range. Skipped cells will not overwrite the existing data of their corresponding cells in the destination range. The default is false.</span></span>

<span data-ttu-id="22038-222">O exemplo de código e as imagens a seguir demonstram esse comportamento em um cenário simples.</span><span class="sxs-lookup"><span data-stu-id="22038-222">The following code sample and images demonstrate this behavior in a simple scenario.</span></span> 

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

<span data-ttu-id="22038-223">*Antes da função anterior ter sido executada.*</span><span class="sxs-lookup"><span data-stu-id="22038-223">*Before the preceeding function has been run.*</span></span>

![Dados no Excel antes do método de cópia do intervalo ter sido executado.](../images/excel-range-copyfrom-skipblanks-before.png)

<span data-ttu-id="22038-225">*Depois que a função anterior foi executada.*</span><span class="sxs-lookup"><span data-stu-id="22038-225">*After the preceeding function has been run.*</span></span>

![Dados no Excel após o método de cópia do intervalo ter sido executado.](../images/excel-range-copyfrom-skipblanks-after.png)

<span data-ttu-id="22038-p110">`transpose` determina se os dados são ou não transpostos, ou seja, suas linhas e colunas são comutadas para o local de origem. Um intervalo transposto é invertido ao longo da diagonal principal, portanto, as linhas  **1**, **2**  e **3** se tornarão as colunas  **A**, **B**  e **C**.</span><span class="sxs-lookup"><span data-stu-id="22038-p110">`transpose` determines whether or not the data is transposed, meaning its rows and columns are switched, into the source location. A transposed range is flipped along the main diagonal, so rows **1**, **2**, and **3** will become columns **A**, **B**, and **C**.</span></span> 


## <a name="see-also"></a><span data-ttu-id="22038-229">Confira também</span><span class="sxs-lookup"><span data-stu-id="22038-229">See also</span></span>

- [<span data-ttu-id="22038-230">Conceitos de programação fundamentais com a API JavaScript do Excel</span><span class="sxs-lookup"><span data-stu-id="22038-230">Fundamental programming concepts with the Excel JavaScript API</span></span>](excel-add-ins-core-concepts.md)

