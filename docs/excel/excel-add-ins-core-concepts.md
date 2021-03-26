---
title: Conceitos fundamentais de programação com a API JavaScript do Excel
description: Use a API JavaScript do Excel para criar suplementos para o Excel.
ms.date: 07/28/2020
localization_priority: Priority
ms.openlocfilehash: dde7dc66e0746fc4d9cf91ed3df824fab05c109d
ms.sourcegitcommit: 9609bd5b4982cdaa2ea7637709a78a45835ffb19
ms.translationtype: HT
ms.contentlocale: pt-BR
ms.lasthandoff: 08/28/2020
ms.locfileid: "47292584"
---
# <a name="fundamental-programming-concepts-with-the-excel-javascript-api"></a><span data-ttu-id="73d62-103">Conceitos fundamentais de programação com a API JavaScript do Excel</span><span class="sxs-lookup"><span data-stu-id="73d62-103">Fundamental programming concepts with the Excel JavaScript API</span></span>

<span data-ttu-id="73d62-104">Este artigo descreve como usar a [API JavaScript do Excel](../reference/overview/excel-add-ins-reference-overview.md) para desenvolver suplementos para o Excel 2016 ou versões posteriores.</span><span class="sxs-lookup"><span data-stu-id="73d62-104">This article describes how to use the [Excel JavaScript API](../reference/overview/excel-add-ins-reference-overview.md) to build add-ins for Excel 2016 or later.</span></span> <span data-ttu-id="73d62-105">Ele apresenta os conceitos básicos que são fundamentais para usar a API e fornece orientações para executar tarefas específicas, como leitura ou gravação em um intervalo grande, atualização de todas as células do intervalo e muito mais.</span><span class="sxs-lookup"><span data-stu-id="73d62-105">It introduces core concepts that are fundamental to using the API and provides guidance for performing specific tasks such as reading or writing to a large range, updating all cells in range, and more.</span></span>

> [!IMPORTANT]
> <span data-ttu-id="73d62-106">Confira [Usar o modelo da API específica do aplicativo](../develop/application-specific-api-model.md) para saber mais sobre a natureza assíncrona das APIs do Excel e como elas funcionam com a pasta de trabalho.</span><span class="sxs-lookup"><span data-stu-id="73d62-106">See [Using the application-specific API model](../develop/application-specific-api-model.md) to learn about the asynchronous nature of the Excel APIs and how they work with the workbook.</span></span>  

## <a name="officejs-apis-for-excel"></a><span data-ttu-id="73d62-107">APIs Office.js para Excel</span><span class="sxs-lookup"><span data-stu-id="73d62-107">Office.js APIs for Excel</span></span>

<span data-ttu-id="73d62-108">Um suplemento do Excel interage com objetos no Excel usando a API JavaScript do Office, que inclui dois modelos de objetos JavaScript:</span><span class="sxs-lookup"><span data-stu-id="73d62-108">An Excel add-in interacts with objects in Excel by using the Office JavaScript API, which includes two JavaScript object models:</span></span>

* <span data-ttu-id="73d62-109">**API JavaScript do Excel**: introduzida com o Office 2016, a [API JavaScript do Excel](../reference/overview/excel-add-ins-reference-overview.md) fornece objetos fortemente tipados que você pode usar para acessar planilhas, intervalos, tabelas, gráficos e muito mais.</span><span class="sxs-lookup"><span data-stu-id="73d62-109">**Excel JavaScript API**: Introduced with Office 2016, the [Excel JavaScript API](../reference/overview/excel-add-ins-reference-overview.md) provides strongly-typed objects that you can use to access worksheets, ranges, tables, charts, and more.</span></span>

* <span data-ttu-id="73d62-110">**APIs Comuns**: Introduzida com o Office 2013, a [API Comum](/javascript/api/office) pode ser usada para acessar recursos como interface de usuário, caixas de diálogo e configurações de cliente, que são comuns entre vários tipos de aplicativos do Office.</span><span class="sxs-lookup"><span data-stu-id="73d62-110">**Common APIs**: Introduced with Office 2013, the [Common API](/javascript/api/office) can be used to access features such as UI, dialogs, and client settings that are common across multiple types of Office applications.</span></span>

<span data-ttu-id="73d62-111">Enquanto você provavelmente use a API JavaScript do Excel para desenvolver a maioria das funcionalidades em suplementos que visam o Excel 2016, você também usará objetos na API comum.</span><span class="sxs-lookup"><span data-stu-id="73d62-111">While you'll likely use the Excel JavaScript API to develop the majority of functionality in add-ins that target Excel 2016 or later, you'll also use objects in the Common API.</span></span> <span data-ttu-id="73d62-112">Por exemplo:</span><span class="sxs-lookup"><span data-stu-id="73d62-112">For example:</span></span>

* <span data-ttu-id="73d62-p103">[Contexto](/javascript/api/office/office.context): o objeto `Context` representa o ambiente de tempo de execução do suplemento e oferece acesso aos principais objetos da API. Ele consiste em detalhes da configuração da pasta de trabalho, como `contentLanguage` e `officeTheme`, além de fornecer informações sobre o ambiente de tempo de execução do suplemento, como `host` e `platform`. Além disso, ele fornece o método `requirements.isSetSupported()`, que você pode usar para verificar se o conjunto de requisitos especificado é suportado pelo aplicativo Excel onde o suplemento está sendo executado.</span><span class="sxs-lookup"><span data-stu-id="73d62-p103">[Context](/javascript/api/office/office.context): The `Context` object represents the runtime environment of the add-in and provides access to key objects of the API. It consists of workbook configuration details such as `contentLanguage` and `officeTheme` and also provides information about the add-in's runtime environment such as `host` and `platform`. Additionally, it provides the `requirements.isSetSupported()` method, which you can use to check whether the specified requirement set is supported by the Excel application where the add-in is running.</span></span>
* <span data-ttu-id="73d62-116">[Documento](/javascript/api/office/office.document): o objeto `Document` fornece o método `getFileAsync()`, que você pode usar para baixar o arquivo do Excel em que o suplemento está sendo executado.</span><span class="sxs-lookup"><span data-stu-id="73d62-116">[Document](/javascript/api/office/office.document): The `Document` object provides the `getFileAsync()` method, which you can use to download the Excel file where the add-in is running.</span></span>

<span data-ttu-id="73d62-117">A imagem a seguir ilustra quando você pode usar a API JavaScript do Excel ou as APIs comuns.</span><span class="sxs-lookup"><span data-stu-id="73d62-117">The following image illustrates when you might use the Excel JavaScript API or the Common APIs.</span></span>

![Imagem das diferentes entre a API JS do Excel e as APIs comuns](../images/excel-js-api-common-api.png)

## <a name="object-model"></a><span data-ttu-id="73d62-119">Modelo de objetos</span><span class="sxs-lookup"><span data-stu-id="73d62-119">Object model</span></span>

<span data-ttu-id="73d62-120">Para entender as APIs do Excel, você deve entender como os componentes de uma pasta de trabalho estão relacionados entre si.</span><span class="sxs-lookup"><span data-stu-id="73d62-120">To understand the Excel APIs, you must understand how the components of a workbook are related to one another.</span></span>

* <span data-ttu-id="73d62-121">Uma **Pasta de trabalho** contém uma ou mais **Planilhas**.</span><span class="sxs-lookup"><span data-stu-id="73d62-121">A **Workbook** contains one or more **Worksheets**.</span></span>
* <span data-ttu-id="73d62-122">Uma **Planilha** concede acesso a células por meio de objetos de **Intervalo**.</span><span class="sxs-lookup"><span data-stu-id="73d62-122">A **Worksheet** gives access to cells through **Range** objects.</span></span>
* <span data-ttu-id="73d62-123">Um **Intervalo** representa um grupo de células contíguas.</span><span class="sxs-lookup"><span data-stu-id="73d62-123">A **Range** represents a group of contiguous cells.</span></span>
* <span data-ttu-id="73d62-124">Os **Intervalos** são usados para criar e colocar **Tabelas**, **Gráficos**, **Formas** e outras visualizações de dados ou objetos da organização.</span><span class="sxs-lookup"><span data-stu-id="73d62-124">**Ranges** are used to create and place **Tables**, **Charts**, **Shapes**, and other data visualization or organization objects.</span></span>
* <span data-ttu-id="73d62-125">Uma **Planilha** contém coleções desses objetos de dados que estão presentes na planilha individual.</span><span class="sxs-lookup"><span data-stu-id="73d62-125">A **Worksheet** contains collections of those data objects that are present in the individual sheet.</span></span>
* <span data-ttu-id="73d62-126">As **Pastas de trabalho** contêm coleções de alguns desses objetos de dados (por exemplo, **Tabelas**) para toda a **Pasta de trabalho**.</span><span class="sxs-lookup"><span data-stu-id="73d62-126">**Workbooks** contain collections of some of those data objects (such as **Tables**) for the entire **Workbook**.</span></span>

### <a name="ranges"></a><span data-ttu-id="73d62-127">Intervalos</span><span class="sxs-lookup"><span data-stu-id="73d62-127">Ranges</span></span>

<span data-ttu-id="73d62-128">Um intervalo é um grupo de células contíguas na pasta de trabalho.</span><span class="sxs-lookup"><span data-stu-id="73d62-128">A range is a group of contiguous cells in the workbook.</span></span> <span data-ttu-id="73d62-129">Os suplementos costumam usar uma notação estilo A1 (por ex.: **B3** para a única célula na coluna **B** e linha **3** ou **C2:F4** para as células das colunas **C** a **F** e linhas **2** a **4**) para definir intervalos.</span><span class="sxs-lookup"><span data-stu-id="73d62-129">Add-ins typically use A1-style notation (e.g. **B3** for the single cell in column **B** and row **3** or **C2:F4** for the cells from columns **C** through **F** and rows **2** through **4**) to define ranges.</span></span>

<span data-ttu-id="73d62-130">Os intervalos têm três propriedades principais: `values`, `formulas` e `format`.</span><span class="sxs-lookup"><span data-stu-id="73d62-130">Ranges have three core properties: `values`, `formulas`, and `format`.</span></span> <span data-ttu-id="73d62-131">Essas propriedades recebem ou definem os valores da célula, as fórmulas a serem avaliadas e a formatação visual das células.</span><span class="sxs-lookup"><span data-stu-id="73d62-131">These properties get or set the cell values, formulas to be evaluated, and the visual formatting of the cells.</span></span>

#### <a name="range-sample"></a><span data-ttu-id="73d62-132">Exemplo de intervalo</span><span class="sxs-lookup"><span data-stu-id="73d62-132">Range sample</span></span>

<span data-ttu-id="73d62-133">O exemplo a seguir mostra como criar registros de vendas.</span><span class="sxs-lookup"><span data-stu-id="73d62-133">The following sample shows how to create sales records.</span></span> <span data-ttu-id="73d62-134">Essa função usa objetos `Range` para definir os valores, fórmulas e formatos.</span><span class="sxs-lookup"><span data-stu-id="73d62-134">This function uses `Range` objects to set the values, formulas, and formats.</span></span>

```js
Excel.run(function (context) {
    var sheet = context.workbook.worksheets.getActiveWorksheet();

    // Create the headers and format them to stand out.
    var headers = [
      ["Product", "Quantity", "Unit Price", "Totals"]
    ];
    var headerRange = sheet.getRange("B2:E2");
    headerRange.values = headers;
    headerRange.format.fill.color = "#4472C4";
    headerRange.format.font.color = "white";

    // Create the product data rows.
    var productData = [
      ["Almonds", 6, 7.5],
      ["Coffee", 20, 34.5],
      ["Chocolate", 10, 9.56],
    ];
    var dataRange = sheet.getRange("B3:D5");
    dataRange.values = productData;

    // Create the formulas to total the amounts sold.
    var totalFormulas = [
      ["=C3 * D3"],
      ["=C4 * D4"],
      ["=C5 * D5"],
      ["=SUM(E3:E5)"]
    ];
    var totalRange = sheet.getRange("E3:E6");
    totalRange.formulas = totalFormulas;
    totalRange.format.font.bold = true;

    // Display the totals as US dollar amounts.
    totalRange.numberFormat = [["$0.00"]];

    return context.sync();
});
```

<span data-ttu-id="73d62-135">Esse exemplo cria os seguintes dados na planilha atual:</span><span class="sxs-lookup"><span data-stu-id="73d62-135">This sample creates the following data in the current worksheet:</span></span>

![Um registro de vendas mostrando as linhas de valores, uma coluna de fórmulas e cabeçalhos formatados.](../images/excel-overview-range-sample.png)

### <a name="charts-tables-and-other-data-objects"></a><span data-ttu-id="73d62-137">Gráficos, tabelas e outros objetos de dados</span><span class="sxs-lookup"><span data-stu-id="73d62-137">Charts, tables, and other data objects</span></span>

<span data-ttu-id="73d62-138">As APIs JavaScript do Excel podem criar e manipular estruturas de dados e visualizações no Excel.</span><span class="sxs-lookup"><span data-stu-id="73d62-138">The Excel JavaScript APIs can create and manipulate the data structures and visualizations within Excel.</span></span> <span data-ttu-id="73d62-139">As tabelas e gráficos são dois dos objetos mais usados, mas as APIs oferecem suporte a tabelas dinâmicas, formas, imagens e muito mais.</span><span class="sxs-lookup"><span data-stu-id="73d62-139">Tables and charts are two of the more commonly used objects, but the APIs support PivotTables, shapes, images, and more.</span></span>

#### <a name="creating-a-table"></a><span data-ttu-id="73d62-140">Criar uma tabela</span><span class="sxs-lookup"><span data-stu-id="73d62-140">Creating a table</span></span>

<span data-ttu-id="73d62-141">Criar tabelas usando intervalos de dados preenchidos.</span><span class="sxs-lookup"><span data-stu-id="73d62-141">Create tables by using data-filled ranges.</span></span> <span data-ttu-id="73d62-142">Controles de formatação e tabela (por exemplo, filtros) são aplicados automaticamente ao intervalo.</span><span class="sxs-lookup"><span data-stu-id="73d62-142">Formatting and table controls (such as filters) are automatically applied to the range.</span></span>

<span data-ttu-id="73d62-143">O exemplo a seguir cria uma tabela usando os intervalos do exemplo anterior.</span><span class="sxs-lookup"><span data-stu-id="73d62-143">The following sample creates a table using the ranges from the previous sample.</span></span>

```js
Excel.run(function (context) {
    var sheet = context.workbook.worksheets.getActiveWorksheet();
    sheet.tables.add("B2:E5", true);
    return context.sync();
});
```

<span data-ttu-id="73d62-144">Usar esse código de exemplo na planilha com os dados anteriores cria a tabela a seguir:</span><span class="sxs-lookup"><span data-stu-id="73d62-144">Using this sample code on the worksheet with the previous data creates the following table:</span></span>

![Uma tabela criada a partir do registro de vendas anterior.](../images/excel-overview-table-sample.png)

#### <a name="creating-a-chart"></a><span data-ttu-id="73d62-146">Criar um gráfico</span><span class="sxs-lookup"><span data-stu-id="73d62-146">Creating a chart</span></span>

<span data-ttu-id="73d62-147">Crie gráficos para visualizar os dados em um intervalo.</span><span class="sxs-lookup"><span data-stu-id="73d62-147">Create charts to visualize the data in a range.</span></span> <span data-ttu-id="73d62-148">As APIs suportam inúmeras variedades de gráficos que podem ser personalizadas de acordo com suas necessidades.</span><span class="sxs-lookup"><span data-stu-id="73d62-148">The APIs support dozens of chart varieties, each of which can be customized to suit your needs.</span></span>

<span data-ttu-id="73d62-149">O exemplo a seguir cria um gráfico de colunas simples para três itens e o coloca 100 pixels abaixo da parte superior da planilha.</span><span class="sxs-lookup"><span data-stu-id="73d62-149">The following sample creates a simple column chart for three items and places it 100 pixels below the top of the worksheet.</span></span>

```js
Excel.run(function (context) {
    var sheet = context.workbook.worksheets.getActiveWorksheet();
    var chart = sheet.charts.add(Excel.ChartType.columnStacked, sheet.getRange("B3:C5"));
    chart.top = 100;
    return context.sync();
});
```

<span data-ttu-id="73d62-150">Executar esse exemplo na planilha com a tabela anterior cria o seguinte gráfico:</span><span class="sxs-lookup"><span data-stu-id="73d62-150">Running this sample on the worksheet with the previous table creates the following chart:</span></span>

![Um gráfico de colunas mostrando as quantidades de três itens do registro de vendas anterior.](../images/excel-overview-chart-sample.png)

## <a name="run-options"></a><span data-ttu-id="73d62-152">Executar opções</span><span class="sxs-lookup"><span data-stu-id="73d62-152">Run options</span></span>

<span data-ttu-id="73d62-153">`Excel.run` tem uma sobrecarga que recebe um objeto [RunOptions](/javascript/api/excel/excel.runoptions).</span><span class="sxs-lookup"><span data-stu-id="73d62-153">`Excel.run` has an overload that takes in a [RunOptions](/javascript/api/excel/excel.runoptions) object.</span></span> <span data-ttu-id="73d62-154">Este contém um conjunto de propriedades que afetam o comportamento de plataforma quando a função é executada.</span><span class="sxs-lookup"><span data-stu-id="73d62-154">This contains a set of properties that affect platform behavior when the function runs.</span></span> <span data-ttu-id="73d62-155">A propriedade a seguir tem suporte no momento:</span><span class="sxs-lookup"><span data-stu-id="73d62-155">The following property is currently supported:</span></span>

* <span data-ttu-id="73d62-156">`delayForCellEdit`: Determina se o Excel atrasa solicitação em lote até que o usuário sai do modo de edição de célula.</span><span class="sxs-lookup"><span data-stu-id="73d62-156">`delayForCellEdit`: Determines whether Excel delays the batch request until the user exits cell edit mode.</span></span> <span data-ttu-id="73d62-157">Quando **verdadeira**, a solicitação em lote é atrasada e executada quando o usuário sai do modo de edição de célula.</span><span class="sxs-lookup"><span data-stu-id="73d62-157">When **true**, the batch request is delayed and runs when the user exits cell edit mode.</span></span> <span data-ttu-id="73d62-158">Quando **falsa**, a solicitação em lote falha automaticamente se o usuário está no modo de edição de célula (causando um erro para alcançar o usuário).</span><span class="sxs-lookup"><span data-stu-id="73d62-158">When **false**, the batch request automatically fails if the user is in cell edit mode (causing an error to reach the user).</span></span> <span data-ttu-id="73d62-159">O comportamento padrão sem nenhuma propriedade `delayForCellEdit` especificada é equivalente a quando é **falsa**.</span><span class="sxs-lookup"><span data-stu-id="73d62-159">The default behavior with no `delayForCellEdit` property specified is equivalent to when it is **false**.</span></span>

```js
Excel.run({ delayForCellEdit: true }, function (context) { ... })
```

## <a name="null-or-blank-property-values"></a><span data-ttu-id="73d62-160">valores de propriedade nulos ou em branco</span><span class="sxs-lookup"><span data-stu-id="73d62-160">null or blank property values</span></span>

<span data-ttu-id="73d62-161">`null` e as cadeias de caracteres esvaziadas têm implicações especiais nas APIs JavaScript do Excel.</span><span class="sxs-lookup"><span data-stu-id="73d62-161">`null` and empty strings have special implications in the Excel JavaScript APIs.</span></span> <span data-ttu-id="73d62-162">Elas são usadas para representar células vazias, sem formatação ou valores padrão.</span><span class="sxs-lookup"><span data-stu-id="73d62-162">They're used to represent empty cells, no formatting, or default values.</span></span> <span data-ttu-id="73d62-163">Essa seção detalha o uso da `null` e de uma cadeia de caracteres vazia ao obter e definir as propriedades.</span><span class="sxs-lookup"><span data-stu-id="73d62-163">This section details the use of `null` and empty string when getting and setting properties.</span></span>

### <a name="null-input-in-2-d-array"></a><span data-ttu-id="73d62-164">entrada nula em uma matriz 2D</span><span class="sxs-lookup"><span data-stu-id="73d62-164">null input in 2-D Array</span></span>

<span data-ttu-id="73d62-p113">No Excel, um intervalo é representado por uma matriz 2D, onde a primeira dimensão é linhas e a segunda dimensão é colunas. Para definir valores, o formato do número ou a fórmula apenas para células específicas em um intervalo, especifique os valores, o formato do número ou a fórmula para essas células na matriz 2D, bem como `null` para todas as outras células na matriz 2D.</span><span class="sxs-lookup"><span data-stu-id="73d62-p113">In Excel, a range is represented by a 2-D array, where the first dimension is rows and the second dimension is columns. To set values, number format, or formula for only specific cells within a range, specify the values, number format, or formula for those cells in the 2-D array, and specify `null` for all other cells in the 2-D array.</span></span>

<span data-ttu-id="73d62-p114">Por exemplo, para atualizar o formato do número apenas para uma célula em um intervalo e manter o formato de número existente para todas as outras células no intervalo, especifique o novo formato de número para a célula a ser atualizada e `null` para todas as outras células. O trecho de código a seguir define um novo formato de número para a quarta célula no intervalo e não altera o formato de número para as primeiras três células no intervalo.</span><span class="sxs-lookup"><span data-stu-id="73d62-p114">For example, to update the number format for only one cell within a range, and retain the existing number format for all other cells in the range, specify the new number format for the cell to update, and specify `null` for all other cells. The following code snippet sets a new number format for the fourth cell in the range, and leaves the number format unchanged for the first three cells in the range.</span></span>

```js
range.values = [['Eurasia', '29.96', '0.25', '15-Feb' ]];
range.numberFormat = [[null, null, null, 'm/d/yyyy;@']];
```

### <a name="null-input-for-a-property"></a><span data-ttu-id="73d62-169">entrada nula para uma propriedade</span><span class="sxs-lookup"><span data-stu-id="73d62-169">null input for a property</span></span>

<span data-ttu-id="73d62-p115">`null` não é uma entrada válida para uma propriedade única. Por exemplo, o trecho de código a seguir não é válido, pois a propriedade `values` do intervalo não pode ser definida como `null`.</span><span class="sxs-lookup"><span data-stu-id="73d62-p115">`null` is not a valid input for single property. For example, the following code snippet is not valid, as the `values` property of the range cannot be set to `null`.</span></span>

```js
range.values = null;
```

<span data-ttu-id="73d62-172">Da mesma forma, o seguinte snippet de código não é válido, pois `null` não é um valor válido para a propriedade `color`.</span><span class="sxs-lookup"><span data-stu-id="73d62-172">Likewise, the following code snippet is not valid, as `null` is not a valid value for the `color` property.</span></span>

```js
range.format.fill.color =  null;
```

### <a name="null-property-values-in-the-response"></a><span data-ttu-id="73d62-173">Valores da propriedade nula na resposta</span><span class="sxs-lookup"><span data-stu-id="73d62-173">null property values in the response</span></span>

<span data-ttu-id="73d62-p116">A formatação de propriedades como `size` e `color` conterá valores `null` na resposta quando valores diferentes existirem no intervalo especificado. Por exemplo, se você recuperar um intervalo e carregar sua propriedade `format.font.color`:</span><span class="sxs-lookup"><span data-stu-id="73d62-p116">Formatting properties such as `size` and `color` will contain `null` values in the response when different values exist in the specified range. For example, if you retrieve a range and load its `format.font.color` property:</span></span>

* <span data-ttu-id="73d62-176">Se todas as células no intervalo tiverem a mesma cor de fonte, `range.format.font.color` especificará essa cor.</span><span class="sxs-lookup"><span data-stu-id="73d62-176">If all cells in the range have the same font color, `range.format.font.color` specifies that color.</span></span>
* <span data-ttu-id="73d62-177">Se houver várias cores de fonte dentro do intervalo, `range.format.font.color` será `null`.</span><span class="sxs-lookup"><span data-stu-id="73d62-177">If multiple font colors are present within the range, `range.format.font.color` is `null`.</span></span>

### <a name="blank-input-for-a-property"></a><span data-ttu-id="73d62-178">Entrada em branco para uma propriedade</span><span class="sxs-lookup"><span data-stu-id="73d62-178">Blank input for a property</span></span>

<span data-ttu-id="73d62-p117">Quando você especificar um valor em branco para uma propriedade (isto é, duas aspas sem espaço entre elas `''`), ele será interpretado como uma instrução para limpar ou redefinir a propriedade. Por exemplo:</span><span class="sxs-lookup"><span data-stu-id="73d62-p117">When you specify a blank value for a property (i.e., two quotation marks with no space in-between `''`), it will be interpreted as an instruction to clear or reset the property. For example:</span></span>

* <span data-ttu-id="73d62-181">Se você especificar um valor em branco para a propriedade `values` de um intervalo, o conteúdo do intervalo será apagado.</span><span class="sxs-lookup"><span data-stu-id="73d62-181">If you specify a blank value for the `values` property of a range, the content of the range is cleared.</span></span>
* <span data-ttu-id="73d62-182">Se você especificar um valor em branco para a propriedade `numberFormat`, o formato de número será redefinido para `General`.</span><span class="sxs-lookup"><span data-stu-id="73d62-182">If you specify a blank value for the `numberFormat` property, the number format is reset to `General`.</span></span>
* <span data-ttu-id="73d62-183">Se você especificar um valor em branco para a propriedade `formula` e a propriedade `formulaLocale`, os valores de fórmula serão apagados.</span><span class="sxs-lookup"><span data-stu-id="73d62-183">If you specify a blank value for the `formula` property and `formulaLocale` property, the formula values are cleared.</span></span>

### <a name="blank-property-values-in-the-response"></a><span data-ttu-id="73d62-184">Valores da propriedade em branco na resposta</span><span class="sxs-lookup"><span data-stu-id="73d62-184">Blank property values in the response</span></span>

<span data-ttu-id="73d62-p118">Para operações de leitura, um valor de propriedade em branco na resposta (isto é, duas aspas sem espaço entre elas `''`) indica que a célula não contém dados nem valor. No primeiro exemplo abaixo, a primeira e a última célula no intervalo não contêm dados. No segundo exemplo, as primeiras duas células no intervalo não contêm uma fórmula.</span><span class="sxs-lookup"><span data-stu-id="73d62-p118">For read operations, a blank property value in the response (i.e., two quotation marks with no space in-between `''`) indicates that cell contains no data or value. In the first example below, the first and last cell in the range contain no data. In the second example, the first two cells in the range do not contain a formula.</span></span>

```js
range.values = [['', 'some', 'data', 'in', 'other', 'cells', '']];
```

```js
range.formula = [['', '', '=Rand()']];
```

## <a name="requirement-sets"></a><span data-ttu-id="73d62-188">Conjuntos de requisitos</span><span class="sxs-lookup"><span data-stu-id="73d62-188">Requirement sets</span></span>

<span data-ttu-id="73d62-189">Os conjuntos de requisitos são grupos nomeados de membros da API.</span><span class="sxs-lookup"><span data-stu-id="73d62-189">Requirement sets are named groups of API members.</span></span> <span data-ttu-id="73d62-190">Um Suplemento do Office pode executar uma verificação de tempo de execução ou usar conjuntos de requisitos especificados no manifesto para determinar se um aplicativo do Office dá suporte às APIs necessárias ao suplemento.</span><span class="sxs-lookup"><span data-stu-id="73d62-190">An Office Add-in can perform a runtime check or use requirement sets specified in the manifest to determine whether an Office application supports the APIs that the add-in needs.</span></span> <span data-ttu-id="73d62-191">Para identificar os conjuntos de requisitos específicos que estão disponíveis em cada plataforma suportada, confira [Conjuntos de requisitos da API JavaScript do Excel](../reference/requirement-sets/excel-api-requirement-sets.md).</span><span class="sxs-lookup"><span data-stu-id="73d62-191">To identify the specific requirement sets that are available on each supported platform, see [Excel JavaScript API requirement sets](../reference/requirement-sets/excel-api-requirement-sets.md).</span></span>

### <a name="checking-for-requirement-set-support-at-runtime"></a><span data-ttu-id="73d62-192">Verificando o suporte ao conjunto de requisitos no tempo de execução</span><span class="sxs-lookup"><span data-stu-id="73d62-192">Checking for requirement set support at runtime</span></span>

<span data-ttu-id="73d62-193">O exemplo de código a seguir mostra como determinar se o aplicativo do Office, onde o suplemento está em execução, dá suporte ao conjunto de requisitos da API especificado.</span><span class="sxs-lookup"><span data-stu-id="73d62-193">The following code sample shows how to determine whether the Office application where the add-in is running supports the specified API requirement set.</span></span>

```js
if (Office.context.requirements.isSetSupported('ExcelApi', '1.3')) {
  /// perform actions
}
else {
  /// provide alternate flow/logic
}
```

### <a name="defining-requirement-set-support-in-the-manifest"></a><span data-ttu-id="73d62-194">Definindo o suporte ao conjunto de requisitos no manifesto</span><span class="sxs-lookup"><span data-stu-id="73d62-194">Defining requirement set support in the manifest</span></span>

<span data-ttu-id="73d62-195">Você pode usar o [elemento Requirements](../reference/manifest/requirements.md) no manifesto do suplemento para especificar os conjuntos de requisitos mínimos e/ou os métodos de API exigidos pelo suplemento para ser ativado.</span><span class="sxs-lookup"><span data-stu-id="73d62-195">You can use the [Requirements element](../reference/manifest/requirements.md) in the add-in manifest to specify the minimal requirement sets and/or API methods that your add-in requires to activate.</span></span> <span data-ttu-id="73d62-196">Se a plataforma ou o aplicativo do Office não der suporte aos conjuntos de requisitos ou aos métodos de API que são especificados no `Requirements`elemento do manifesto, o suplemento não será executado nesse aplicativo ou plataforma e não será exibido na lista de suplementos que são mostrados em **Meus Suplementos**.</span><span class="sxs-lookup"><span data-stu-id="73d62-196">If the Office application or platform doesn't support the requirement sets or API methods that are specified in the `Requirements` element of the manifest, the add-in won't run in that application or platform, and it won't display in the list of add-ins that are shown in **My Add-ins**.</span></span>

<span data-ttu-id="73d62-197">O exemplo de código a seguir mostra o elemento `Requirements` em um manifesto de suplemento que especifica se o suplemento deve ser carregado em todos os aplicativos cliente do Office que dão suporte ao conjunto de requisitos ExcelApi, versão 1.3 ou superior.</span><span class="sxs-lookup"><span data-stu-id="73d62-197">The following code sample shows the `Requirements` element in an add-in manifest which specifies that the add-in should load in all Office client applications that support ExcelApi requirement set version 1.3 or greater.</span></span>

```xml
<Requirements>
   <Sets DefaultMinVersion="1.3">
      <Set Name="ExcelApi" MinVersion="1.3"/>
   </Sets>
</Requirements>
```

> [!NOTE]
> <span data-ttu-id="73d62-198">Para disponibilizar seu suplemento em todas as plataformas de um aplicativo do Office, como Excel Online, Windows e iPad, é recomendável verificar o suporte a requisitos no tempo de execução, em vez de definir o suporte ao conjunto de requisitos no manifesto.</span><span class="sxs-lookup"><span data-stu-id="73d62-198">To make your add-in available on all platforms of an Office application, such as Excel on the web, Windows, and iPad, we recommend that you check for requirement support at runtime instead of defining requirement set support in the manifest.</span></span>

### <a name="requirement-sets-for-the-officejs-common-api"></a><span data-ttu-id="73d62-199">Conjuntos de requisitos para a API comum Office.js</span><span class="sxs-lookup"><span data-stu-id="73d62-199">Requirement sets for the Office.js Common API</span></span>

<span data-ttu-id="73d62-200">Para saber mais sobre conjuntos de requisitos comuns da API, confira [Conjuntos de requisitos comuns da API do Office](../reference/requirement-sets/office-add-in-requirement-sets.md).</span><span class="sxs-lookup"><span data-stu-id="73d62-200">For information about Common API requirement sets, see [Office Common API requirement sets](../reference/requirement-sets/office-add-in-requirement-sets.md).</span></span>

## <a name="handle-errors"></a><span data-ttu-id="73d62-201">Lidar com erros</span><span class="sxs-lookup"><span data-stu-id="73d62-201">Handle errors</span></span>

<span data-ttu-id="73d62-202">Quando ocorre um erro de API, a API retorna um objeto `error` que contém um código e uma mensagem.</span><span class="sxs-lookup"><span data-stu-id="73d62-202">When an API error occurs, the API returns an `error` object that contains a code and a message.</span></span> <span data-ttu-id="73d62-203">Para saber mais sobre o tratamento de erros, incluindo uma lista de erros da API, confira [Tratamento de erro](excel-add-ins-error-handling.md).</span><span class="sxs-lookup"><span data-stu-id="73d62-203">For detailed information about error handling, including a list of API errors, see [Error handling](excel-add-ins-error-handling.md).</span></span>

## <a name="see-also"></a><span data-ttu-id="73d62-204">Confira também</span><span class="sxs-lookup"><span data-stu-id="73d62-204">See also</span></span>

* [<span data-ttu-id="73d62-205">Crie seu primeiro suplemento do Excel</span><span class="sxs-lookup"><span data-stu-id="73d62-205">Build your first Excel add-in</span></span>](../quickstarts/excel-quickstart-jquery.md)
* [<span data-ttu-id="73d62-206">Exemplos de código de suplementos do Excel</span><span class="sxs-lookup"><span data-stu-id="73d62-206">Excel add-ins code samples</span></span>](https://developer.microsoft.com/office/gallery/?filterBy=Samples,Excel)
* [<span data-ttu-id="73d62-207">Otimização de desempenho da API JavaScript do Excel</span><span class="sxs-lookup"><span data-stu-id="73d62-207">Excel JavaScript API performance optimization</span></span>](../excel/performance.md)
* [<span data-ttu-id="73d62-208">Referência da API JavaScript do Excel</span><span class="sxs-lookup"><span data-stu-id="73d62-208">Excel JavaScript API reference</span></span>](../reference/overview/excel-add-ins-reference-overview.md)
