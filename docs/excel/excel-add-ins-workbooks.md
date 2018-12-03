---
title: Trabalhar com pastas de trabalho usando a API JavaScript do Excel
description: ''
ms.date: 11/27/2018
ms.openlocfilehash: 1cfde9bfdf306e35f47595f936679d9fa6e1814e
ms.sourcegitcommit: 026437bd3819f4e9cd4153ebe60c98ab04e18f4e
ms.translationtype: HT
ms.contentlocale: pt-BR
ms.lasthandoff: 11/30/2018
ms.locfileid: "27002334"
---
# <a name="work-with-workbooks-using-the-excel-javascript-api"></a><span data-ttu-id="54917-102">Trabalhar com pastas de trabalho usando a API JavaScript do Excel</span><span class="sxs-lookup"><span data-stu-id="54917-102">Work with Charts using the Excel JavaScript API</span></span>

<span data-ttu-id="54917-103">Este artigo fornece exemplos de código que mostram como executar tarefas comuns com pastas de trabalho usando a API JavaScript do Excel.</span><span class="sxs-lookup"><span data-stu-id="54917-103">This article provides code samples that show how to perform common tasks with ranges using the Excel JavaScript API.</span></span> <span data-ttu-id="54917-104">Para obter a lista completa de propriedades e métodos que o objeto **Workbook** suporta, confira [Objeto Workbook (API JavaScript para Excel)](/javascript/api/excel/excel.workbook).</span><span class="sxs-lookup"><span data-stu-id="54917-104">For the complete list of properties and methods that the **Range** object supports, see [Range Object (JavaScript API for Excel)](/javascript/api/excel/excel.workbook).</span></span> <span data-ttu-id="54917-105">Este artigo aborda também ações em nível de pasta de trabalho executadas através do objeto [Application](/javascript/api/excel/excel.application).</span><span class="sxs-lookup"><span data-stu-id="54917-105">This article also covers workbook-level actions performed through the [Application](/javascript/api/excel/excel.application) object.</span></span>

<span data-ttu-id="54917-106">O objeto Workbook é o ponto de entrada para que se suplemento interaja com o Excel.</span><span class="sxs-lookup"><span data-stu-id="54917-106">The Workbook object is the entry point for your add-in to interact with Excel.</span></span> <span data-ttu-id="54917-107">Ele mantém conjuntos de planilhas, tabelas, Tabelas Dinâmicas e muito mais, através dos quais os dados do Excel são acessados e alterados.</span><span class="sxs-lookup"><span data-stu-id="54917-107">It maintains collections of worksheets, tables, PivotTables, and more, through which Excel data is accessed and changed.</span></span> <span data-ttu-id="54917-108">O objeto [WorksheetCollection](/javascript/api/excel/excel.worksheetcollection) dá a seu suplemento acesso a todos os dados de pastas de trabalho através de planilhas individuais.</span><span class="sxs-lookup"><span data-stu-id="54917-108">The [WorksheetCollection](/javascript/api/excel/excel.worksheetcollection) object gives your add-in access to all the workbook's data through indivual worksheets.</span></span> <span data-ttu-id="54917-109">Especificamente, ele permite seu suplemento adicione planilhas, navegue entre elas e atribua manipuladores a eventos de planilhas.</span><span class="sxs-lookup"><span data-stu-id="54917-109">Specifically, it lets your add-in add worksheets, navigate among them, and assign handlers to worksheet events.</span></span> <span data-ttu-id="54917-110">O artigo [Trabalhar com planilhas usando a API JavaScript do Excel](excel-add-ins-worksheets.md) descreve como acessar e editar planilhas.</span><span class="sxs-lookup"><span data-stu-id="54917-110">The article [Work with worksheets using the Excel JavaScript API](excel-add-ins-worksheets.md) describes how to access and edit worksheets.</span></span>

## <a name="get-the-active-cell-or-selected-range"></a><span data-ttu-id="54917-111">Obter a célula ativa ou o intervalo selecionado</span><span class="sxs-lookup"><span data-stu-id="54917-111">Get the active cell or selected range</span></span>

<span data-ttu-id="54917-112">O objeto Workbook contém dois métodos que obtêm um intervalo de células que o usuário ou o suplemento selecionaram: `getActiveCell()` e `getSelectedRange()`.</span><span class="sxs-lookup"><span data-stu-id="54917-112">The Workbook object contains two methods that get a range of cells the user or add-in has selected: `getActiveCell()` and `getSelectedRange()`.</span></span> <span data-ttu-id="54917-113">`getActiveCell()` obtém a célula ativa da pasta de trabalho como um [objeto Range](/javascript/api/excel/excel.range).</span><span class="sxs-lookup"><span data-stu-id="54917-113">`getActiveCell()` gets the active cell from the workbook as a [Range object](/javascript/api/excel/excel.range).</span></span> <span data-ttu-id="54917-114">O exemplo a seguir mostra uma chamada para `getActiveCell()`, seguida do endereço da célula que está sendo impresso no console.</span><span class="sxs-lookup"><span data-stu-id="54917-114">The following example shows a call to `getActiveCell()`, followed by the cell's address being printed to the console.</span></span>

```js
Excel.run(function (context) {
    var activeCell = context.workbook.getActiveCell();
    activeCell.load("address");

    return context.sync().then(function () {
        console.log("The active cell is " + activeCell.address);
    });
}).catch(errorHandlerFunction);
```

<span data-ttu-id="54917-115">O método `getSelectedRange()` retorna o intervalo único selecionado atualmente.</span><span class="sxs-lookup"><span data-stu-id="54917-115">The `getSelectedRange()` method returns the currently selected single range.</span></span> <span data-ttu-id="54917-116">Se houver vários intervalos selecionados, será gerado um erro InvalidSelection.</span><span class="sxs-lookup"><span data-stu-id="54917-116">If multiple ranges are selected, an InvalidSelection error is thrown.</span></span> <span data-ttu-id="54917-117">O exemplo a seguir mostra uma chamada para `getSelectedRange()` que, em seguida, define a cor de preenchimento do intervalo como amarelo.</span><span class="sxs-lookup"><span data-stu-id="54917-117">The following example shows a call to `getSelectedRange()` that then sets the range's fill color to yellow.</span></span>

```js
Excel.run(function(context) {
    var range = context.workbook.getSelectedRange();
    range.format.fill.color = "yellow";
    return context.sync();
}).catch(errorHandlerFunction);
```

## <a name="create-a-workbook"></a><span data-ttu-id="54917-118">Criar uma pasta de trabalho</span><span class="sxs-lookup"><span data-stu-id="54917-118">Create a workbook</span></span>

<span data-ttu-id="54917-119">O suplemento pode criar uma nova pasta de trabalho separada da instância do Excel, na qual o suplemento está sendo executado atualmente.</span><span class="sxs-lookup"><span data-stu-id="54917-119">Your add-in can create a new workbook, separate from the Excel instance in which the add-in is currently running.</span></span> <span data-ttu-id="54917-120">O objeto do Excel tem o método `createWorkbook` para esta finalidade.</span><span class="sxs-lookup"><span data-stu-id="54917-120">The Excel object has the `createWorkbook` method for this purpose.</span></span> <span data-ttu-id="54917-121">Quando esse método é chamado, a nova pasta de trabalho é aberta imediatamente e exibida em uma nova instância do Excel.</span><span class="sxs-lookup"><span data-stu-id="54917-121">When this method is called, the new workbook is immediately opened and displayed in a new instance of Excel.</span></span> <span data-ttu-id="54917-122">O suplemento permanece aberto e em execução com a pasta de trabalho anterior.</span><span class="sxs-lookup"><span data-stu-id="54917-122">Your add-in remains open and running with the previous workbook.</span></span>

```js
Excel.createWorkbook();
```

<span data-ttu-id="54917-123">O método `createWorkbook` também cria uma cópia de uma pasta de trabalho existente.</span><span class="sxs-lookup"><span data-stu-id="54917-123">The `createWorkbook` method can also create a copy of an existing workbook.</span></span> <span data-ttu-id="54917-124">O método aceita uma representação de cadeia de caracteres codificada em Base64 de um arquivo .xlsx como parâmetro opcional.</span><span class="sxs-lookup"><span data-stu-id="54917-124">The method accepts a base64-encoded string representation of an .xlsx file as an optional parameter.</span></span> <span data-ttu-id="54917-125">A pasta de trabalho resultante será uma cópia desse arquivo, supondo que o argumento da cadeia de caracteres seja um arquivo. xlsx válido.</span><span class="sxs-lookup"><span data-stu-id="54917-125">The resulting workbook will be a copy of that file, assuming the string argument is a valid .xlsx file.</span></span>

<span data-ttu-id="54917-126">Você pode obter a pasta de trabalho atual do suplemento como uma cadeia de caracteres codificada com Base64 usando a [divisão de arquivos](/javascript/api/office/office.document#getfileasync-filetype--options--callback-).</span><span class="sxs-lookup"><span data-stu-id="54917-126">You can get your add-in’s current workbook as a base64-encoded string by using [file slicing](/javascript/api/office/office.document#getfileasync-filetype--options--callback-).</span></span> <span data-ttu-id="54917-127">A classe [FileReader](https://developer.mozilla.org/docs/Web/API/FileReader) pode ser usada para converter um arquivo em uma cadeia de caracteres codificada com Base64, como demonstrado no seguinte exemplo.</span><span class="sxs-lookup"><span data-stu-id="54917-127">The [FileReader](https://developer.mozilla.org/docs/Web/API/FileReader) class can be used to convert a file into the required base64-encoded string, as demonstrated in the following example.</span></span> 

```js
var myFile = document.getElementById("file");
var reader = new FileReader();

reader.onload = (function (event) {
    Excel.run(function (context) {
        // strip off the metadata before the base64-encoded string
        var startIndex = event.target.result.indexOf("base64,");
        var mybase64 = event.target.result.substr(startIndex + 7);

        Excel.createWorkbook(mybase64);
        return context.sync();
    }).catch(errorHandlerFunction);
});

// read in the file as a data URL so we can parse the base64-encoded string
reader.readAsDataURL(myFile.files[0]);
```

## <a name="protect-the-workbooks-structure"></a><span data-ttu-id="54917-128">Proteger a estrutura da pasta de trabalho</span><span class="sxs-lookup"><span data-stu-id="54917-128">Protect the workbook's structure</span></span>

<span data-ttu-id="54917-129">O suplemento pode controlar a capacidade de um usuário de editar dados em uma pasta de trabalho.</span><span class="sxs-lookup"><span data-stu-id="54917-129">Your add-in can control a user's ability to edit the workbook's structure.</span></span> <span data-ttu-id="54917-130">A propriedade `protection` do objeto Workbook é um objeto [WorkbookProtection](/javascript/api/excel/excel.workbookprotection) com um método `protect()`.</span><span class="sxs-lookup"><span data-stu-id="54917-130">The Workbook object's `protection` property is a [WorkbookProtection](/javascript/api/excel/excel.workbookprotection) object with a `protect()` method.</span></span> <span data-ttu-id="54917-131">O exemplo a seguir mostra um cenário básico ativando/desativando a proteção da pasta de trabalho.</span><span class="sxs-lookup"><span data-stu-id="54917-131">The following example shows a basic scenario toggling the protection of the workbook's structure.</span></span> 

```js
Excel.run(function (context) {
    var workbook = context.workbook;
    workbook.load("protection/protected");

    return context.sync().then(function() {
        if (!workbook.protection.protected) {
            workbook.protection.protect();
        }
    });
}).catch(errorHandlerFunction);
```

<span data-ttu-id="54917-132">O método `protect` aceita um parâmetro opcional de cadeia de caracteres.</span><span class="sxs-lookup"><span data-stu-id="54917-132">The `protect` method accepts an optional string parameter.</span></span> <span data-ttu-id="54917-133">Esta cadeia de caracteres representa a senha necessária para um usuário ignorar a proteção e alterar a estrutura da pasta de trabalho.</span><span class="sxs-lookup"><span data-stu-id="54917-133">This string represents the password needed for a user to bypass protection and change the workbook's structure.</span></span>

<span data-ttu-id="54917-134">A proteção também ser definida no nível da planilha para prevenir a edição de dados indesejada.</span><span class="sxs-lookup"><span data-stu-id="54917-134">Protection can also be set at the worksheet level to prevent unwanted data editing.</span></span> <span data-ttu-id="54917-135">Para saber mais, confira a seção **Proteção de dados**do artigo [Trabalhar com planilhas usando a API JavaScript do Excel](excel-add-ins-worksheets.md#data-protection).</span><span class="sxs-lookup"><span data-stu-id="54917-135">For more information, see the **Data protection** section of the [Work with worksheets using the Excel JavaScript API](excel-add-ins-worksheets.md#data-protection) article.</span></span>

> [!NOTE] 
> <span data-ttu-id="54917-136">Para saber mais sobre a proteção de pastas de trabalho no Excel, confira o artigo [Proteger uma pasta de trabalho](https://support.office.com/article/Protect-a-workbook-7E365A4D-3E89-4616-84CA-1931257C1517).</span><span class="sxs-lookup"><span data-stu-id="54917-136">For more information about workbook protection in Excel, see the [Protect a workbook](https://support.office.com/article/Protect-a-workbook-7E365A4D-3E89-4616-84CA-1931257C1517) article.</span></span>

## <a name="access-document-properties"></a><span data-ttu-id="54917-137">Acessar propriedades do documentos</span><span class="sxs-lookup"><span data-stu-id="54917-137">Access document properties</span></span>

<span data-ttu-id="54917-138">Objetos Workbook têm acesso aos metadados dos arquivos do Office, que são conhecidos como [propriedades de documentos](https://support.office.com/article/View-or-change-the-properties-for-an-Office-file-21D604C2-481E-4379-8E54-1DD4622C6B75).</span><span class="sxs-lookup"><span data-stu-id="54917-138">Workbook objects have access to the Office file metadata, which is known as the [document properties](https://support.office.com/article/View-or-change-the-properties-for-an-Office-file-21D604C2-481E-4379-8E54-1DD4622C6B75).</span></span> <span data-ttu-id="54917-139">A propriedade `properties` do objeto Workbook é um objeto [DocumentProperties](/javascript/api/excel/excel.documentproperties) que contém esses valores de metadados.</span><span class="sxs-lookup"><span data-stu-id="54917-139">The Workbook object's `properties` property is a [DocumentProperties](/javascript/api/excel/excel.documentproperties) object containing these metadata values.</span></span> <span data-ttu-id="54917-140">O exemplo a seguir mostra como definir a propriedade **author**.</span><span class="sxs-lookup"><span data-stu-id="54917-140">The following example shows how to set the **MetadataCatalogFileName** property declaratively.</span></span>

```js
Excel.run(function (context) {
    var docProperties = context.workbook.properties;
    docProperties.author = "Alex";
    return context.sync();
}).catch(errorHandlerFunction);
```

<span data-ttu-id="54917-141">Você também pode definir propriedades personalizadas.</span><span class="sxs-lookup"><span data-stu-id="54917-141">You can also define custom properties.</span></span> <span data-ttu-id="54917-142">O objeto DocumentProperties contém uma propriedade `custom` que representa um conjunto de pares de valores-chave para propriedades definidas pelo usuário.</span><span class="sxs-lookup"><span data-stu-id="54917-142">The DocumentProperties object contains a `custom` property that represents a collection of key-value pairs for user-defined properties.</span></span> <span data-ttu-id="54917-143">O exemplo a seguir mostra como criar uma propriedade personalizada chamada **Introduction** com o valor "Olá" e, em seguida, recuperá-la.</span><span class="sxs-lookup"><span data-stu-id="54917-143">The following example shows how to create a custom property named **Introduction** with the value "Hello", then retrieve it.</span></span>

```js
Excel.run(function (context) {
    var customDocProperties = context.workbook.properties.custom;
    customDocProperties.add("Introduction", "Hello");
    return context.sync();
}).catch(errorHandlerFunction);

[...]

Excel.run(function (context) {
    var customDocProperties = context.workbook.properties.custom;
    var customProperty = customDocProperties.getItem("Introduction");
    customProperty.load("key, value");

    return context.sync().then(function() {
        console.log("Custom key  : " + customProperty.key); // "Introduction"
        console.log("Custom value : " + customProperty.value); // "Hello"
    });
}).catch(errorHandlerFunction);
```

## <a name="access-document-settings"></a><span data-ttu-id="54917-144">Acessar configurações do documentos</span><span class="sxs-lookup"><span data-stu-id="54917-144">Access document settings</span></span>

<span data-ttu-id="54917-145">As configurações da pasta de trabalho são semelhantes ao conjunto de propriedades personalizadas.</span><span class="sxs-lookup"><span data-stu-id="54917-145">A workbook's settings are similar to the collection of custom properties.</span></span> <span data-ttu-id="54917-146">A diferença é que as configurações são exclusivas para um único arquivo do Excel e emparelhamento de suplementos, enquanto que as propriedades estão somente conectadas ao arquivo.</span><span class="sxs-lookup"><span data-stu-id="54917-146">The difference is settings are unique to a single Excel file and add-in pairing, whereas properties are solely connected to the file.</span></span> <span data-ttu-id="54917-147">O exemplo a seguir mostra como criar e acessar uma configuração.</span><span class="sxs-lookup"><span data-stu-id="54917-147">The following example shows how to  create a file and add it to a folder.</span></span>

```js
Excel.run(function (context) {
    var settings = context.workbook.settings;
    settings.add("NeedsReview", true);
    var needsReview = settings.getItem("NeedsReview");
    needsReview.load("value");

    return context.sync().then(function() {
        console.log("Workbook needs review : " + needsReview.value);
    });
}).catch(errorHandlerFunction);
```

## <a name="control-calculation-behavior"></a><span data-ttu-id="54917-148">Controlar o comportamento do cálculo</span><span class="sxs-lookup"><span data-stu-id="54917-148">Control calculation behavior</span></span>

### <a name="set-calculation-mode"></a><span data-ttu-id="54917-149">Configurar o modo de cálculo</span><span class="sxs-lookup"><span data-stu-id="54917-149">Set calculation mode</span></span>

<span data-ttu-id="54917-150">Por padrão, o Excel recalcula os resultados das fórmulas sempre que uma célula referenciada é alterada.</span><span class="sxs-lookup"><span data-stu-id="54917-150">By default, Excel recalculates formula results whenever a referenced cell is changed.</span></span> <span data-ttu-id="54917-151">O desempenho de seu suplemento pode se beneficiar do ajuste desse comportamento de cálculo.</span><span class="sxs-lookup"><span data-stu-id="54917-151">Your add-in's performance may benefit from adjusting this calculation behavior.</span></span> <span data-ttu-id="54917-152">O objeto Application tem uma propriedade `calculationMode` do tipo `CalculationMode`.</span><span class="sxs-lookup"><span data-stu-id="54917-152">The Application object has a `calculationMode` property of type `CalculationMode`.</span></span> <span data-ttu-id="54917-153">Esta propriedade pode ser configurada com os seguintes valores:</span><span class="sxs-lookup"><span data-stu-id="54917-153">It can be set to the following values:</span></span>

 - <span data-ttu-id="54917-154">`automatic`: O comportamento de recálculo padrão em que o Excel calcula novos resultados das fórmulas sempre que o dado relevante é alterado.</span><span class="sxs-lookup"><span data-stu-id="54917-154">`automatic`: The default recalculation behavior where Excel calculates new formula results every time the relevant data is changed.</span></span>
 - <span data-ttu-id="54917-155">`automaticExceptTables`: Igual a `automatic`, exceto que as alterações feitas nos valores em tabelas serão ignoradas.</span><span class="sxs-lookup"><span data-stu-id="54917-155">`automaticExceptTables`: Same as `automatic`, except any changes made to values in tables are ignored.</span></span>
 - <span data-ttu-id="54917-156">`manual`: Os cálculos ocorrem somente quando o usuário ou suplemento os solicita.</span><span class="sxs-lookup"><span data-stu-id="54917-156">`manual`: Calculations only occur when the user or add-in requests them.</span></span>
 
### <a name="set-calculation-type"></a><span data-ttu-id="54917-157">Configurar o tipo de cálculo</span><span class="sxs-lookup"><span data-stu-id="54917-157">Set calculation type</span></span>

<span data-ttu-id="54917-158">O objeto [Application](/javascript/api/excel/excel.application) fornece um método para forçar um recálculo imediato.</span><span class="sxs-lookup"><span data-stu-id="54917-158">The [Application](/javascript/api/excel/excel.application) object provides a method to force an immediate recalculation.</span></span> <span data-ttu-id="54917-159">`Application.calculate(calculationType)` inicia o recálculo manual baseado no `calculationType` especificado.</span><span class="sxs-lookup"><span data-stu-id="54917-159">`Application.calculate(calculationType)` starts a manual recalculation based on the specified `calculationType`.</span></span> <span data-ttu-id="54917-160">Os seguintes valores podem ser especificados:</span><span class="sxs-lookup"><span data-stu-id="54917-160">The following values can be specified:</span></span>

 - <span data-ttu-id="54917-161">`full`: Recalcule todas as fórmulas em todas as pastas de trabalho abertas, independentemente de elas terem sido alteradas desde o último recálculo.</span><span class="sxs-lookup"><span data-stu-id="54917-161">`full`: Recalculate all formulas in all open workbooks, regardless of whether they have changed since the last recalculation.</span></span>
 - <span data-ttu-id="54917-162">`fullRebuild`: Verifique as fórmulas dependentes e depois recalcule todas as fórmulas em todas as pastas de trabalho abertas, independentemente de elas terem sido alteradas desde o último recálculo.</span><span class="sxs-lookup"><span data-stu-id="54917-162">`fullRebuild`: Check dependent formulas, and then recalculate all formulas in all open workbooks, regardless of whether they have changed since the last recalculation.</span></span>
 - <span data-ttu-id="54917-163">`recalculate`: Recalcule as fórmulas que foram alteradas (ou marcadas por programação para recálculo) desde o último cálculo, e as fórmulas dependentes nelas, em todas as pastas de trabalho ativas.</span><span class="sxs-lookup"><span data-stu-id="54917-163">`recalculate`: Recalculate formulas that have changed (or been programmatically marked for recalculation) since the last calculation, and formulas dependent on them, in all active workbooks.</span></span>
 
> [!NOTE] 
> <span data-ttu-id="54917-164">Para saber mais sobre o recálculo, confira o artigo [Alterar o recálculo, a iteração ou a precisão de fórmulas](https://support.office.com/article/change-formula-recalculation-iteration-or-precision-73fc7dac-91cf-4d36-86e8-67124f6bcce4).</span><span class="sxs-lookup"><span data-stu-id="54917-164">For more information about recalculation, see the [Change formula recalculation, iteration, or precision](https://support.office.com/article/change-formula-recalculation-iteration-or-precision-73fc7dac-91cf-4d36-86e8-67124f6bcce4) article.</span></span>

### <a name="temporarily-suspend-calculations"></a><span data-ttu-id="54917-165">Suspender os cálculos temporariamente</span><span class="sxs-lookup"><span data-stu-id="54917-165">Temporarily suspend calculations</span></span>

<span data-ttu-id="54917-166">A API do Excel também permite que os suplementos desativem os cálculos até que `RequestContext.sync()` seja chamado.</span><span class="sxs-lookup"><span data-stu-id="54917-166">The Excel API also lets add-ins turn off calculations until `RequestContext.sync()` is called.</span></span> <span data-ttu-id="54917-167">Isso é feito pelo `suspendApiCalculationUntilNextSync()`.</span><span class="sxs-lookup"><span data-stu-id="54917-167">This is done with `suspendApiCalculationUntilNextSync()`.</span></span> <span data-ttu-id="54917-168">Use esse método quando seu suplemento estiver editando intervalos extensos sem precisar acessar os dados entre as edições.</span><span class="sxs-lookup"><span data-stu-id="54917-168">Use this method when your add-in is editing large ranges without needing to access the data between edits.</span></span>

```js
context.application.suspendApiCalculationUntilNextSync();
```

## <a name="see-also"></a><span data-ttu-id="54917-169">Confira também</span><span class="sxs-lookup"><span data-stu-id="54917-169">See also</span></span>

- [<span data-ttu-id="54917-170">Conceitos fundamentais de programação com a API JavaScript do Excel</span><span class="sxs-lookup"><span data-stu-id="54917-170">Fundamental programming concepts with the Excel JavaScript API</span></span>](excel-add-ins-core-concepts.md)
- [<span data-ttu-id="54917-171">Trabalhar com planilhas usando a API JavaScript do Excel</span><span class="sxs-lookup"><span data-stu-id="54917-171">Work with Worksheets using the Excel JavaScript API</span></span>](excel-add-ins-worksheets.md)
- [<span data-ttu-id="54917-172">Trabalhar com intervalos usando a API JavaScript do Excel</span><span class="sxs-lookup"><span data-stu-id="54917-172">Work with ranges using the Excel JavaScript API</span></span>](excel-add-ins-ranges.md)