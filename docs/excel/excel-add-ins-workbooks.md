---
title: Trabalhar com pastas de trabalho usando a API JavaScript do Excel
description: Saiba como executar tarefas comuns com as guias de trabalho ou recursos no nível do aplicativo usando Excel API JavaScript.
ms.date: 06/07/2021
ms.prod: excel
localization_priority: Normal
ms.openlocfilehash: 78cdf57ce6ecce3e9e3e40188b3325cdf15ab265
ms.sourcegitcommit: 883f71d395b19ccfc6874a0d5942a7016eb49e2c
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 07/09/2021
ms.locfileid: "53349424"
---
# <a name="work-with-workbooks-using-the-excel-javascript-api"></a><span data-ttu-id="52707-103">Trabalhar com pastas de trabalho usando a API JavaScript do Excel</span><span class="sxs-lookup"><span data-stu-id="52707-103">Work with workbooks using the Excel JavaScript API</span></span>

<span data-ttu-id="52707-104">Este artigo fornece exemplos de código que mostram como executar tarefas comuns com pastas de trabalho usando a API JavaScript do Excel.</span><span class="sxs-lookup"><span data-stu-id="52707-104">This article provides code samples that show how to perform common tasks with workbooks using the Excel JavaScript API.</span></span> <span data-ttu-id="52707-105">Para ver a lista completa de propriedades e métodos que o objeto oferece suporte, consulte `Workbook` [Objeto Workbook (API JavaScript para Excel)](/javascript/api/excel/excel.workbook).</span><span class="sxs-lookup"><span data-stu-id="52707-105">For the complete list of properties and methods that the `Workbook` object supports, see [Workbook Object (JavaScript API for Excel)](/javascript/api/excel/excel.workbook).</span></span> <span data-ttu-id="52707-106">Este artigo aborda também ações em nível de pasta de trabalho executadas através do objeto [Application](/javascript/api/excel/excel.application).</span><span class="sxs-lookup"><span data-stu-id="52707-106">This article also covers workbook-level actions performed through the [Application](/javascript/api/excel/excel.application) object.</span></span>

<span data-ttu-id="52707-107">O objeto Workbook é o ponto de entrada para que se suplemento interaja com o Excel.</span><span class="sxs-lookup"><span data-stu-id="52707-107">The Workbook object is the entry point for your add-in to interact with Excel.</span></span> <span data-ttu-id="52707-108">Ele mantém conjuntos de planilhas, tabelas, Tabelas Dinâmicas e muito mais, através dos quais os dados do Excel são acessados e alterados.</span><span class="sxs-lookup"><span data-stu-id="52707-108">It maintains collections of worksheets, tables, PivotTables, and more, through which Excel data is accessed and changed.</span></span> <span data-ttu-id="52707-109">O objeto [WorksheetCollection](/javascript/api/excel/excel.worksheetcollection) dá a seu suplemento acesso a todos os dados de pastas de trabalho através de planilhas individuais.</span><span class="sxs-lookup"><span data-stu-id="52707-109">The [WorksheetCollection](/javascript/api/excel/excel.worksheetcollection) object gives your add-in access to all the workbook's data through individual worksheets.</span></span> <span data-ttu-id="52707-110">Especificamente, ele permite seu suplemento adicione planilhas, navegue entre elas e atribua manipuladores a eventos de planilhas.</span><span class="sxs-lookup"><span data-stu-id="52707-110">Specifically, it lets your add-in add worksheets, navigate among them, and assign handlers to worksheet events.</span></span> <span data-ttu-id="52707-111">O artigo [Trabalhar com planilhas usando a API JavaScript do Excel](excel-add-ins-worksheets.md) descreve como acessar e editar planilhas.</span><span class="sxs-lookup"><span data-stu-id="52707-111">The article [Work with worksheets using the Excel JavaScript API](excel-add-ins-worksheets.md) describes how to access and edit worksheets.</span></span>

## <a name="get-the-active-cell-or-selected-range"></a><span data-ttu-id="52707-112">Obter a célula ativa ou o intervalo selecionado</span><span class="sxs-lookup"><span data-stu-id="52707-112">Get the active cell or selected range</span></span>

<span data-ttu-id="52707-113">O objeto Workbook contém dois métodos que obtêm um intervalo de células que o usuário ou o suplemento selecionaram: `getActiveCell()` e `getSelectedRange()`.</span><span class="sxs-lookup"><span data-stu-id="52707-113">The Workbook object contains two methods that get a range of cells the user or add-in has selected: `getActiveCell()` and `getSelectedRange()`.</span></span> <span data-ttu-id="52707-114">`getActiveCell()` obtém a célula ativa da pasta de trabalho como um [objeto Range](/javascript/api/excel/excel.range).</span><span class="sxs-lookup"><span data-stu-id="52707-114">`getActiveCell()` gets the active cell from the workbook as a [Range object](/javascript/api/excel/excel.range).</span></span> <span data-ttu-id="52707-115">O exemplo a seguir mostra uma chamada para `getActiveCell()`, seguida do endereço da célula que está sendo impresso no console.</span><span class="sxs-lookup"><span data-stu-id="52707-115">The following example shows a call to `getActiveCell()`, followed by the cell's address being printed to the console.</span></span>

```js
Excel.run(function (context) {
    var activeCell = context.workbook.getActiveCell();
    activeCell.load("address");

    return context.sync().then(function () {
        console.log("The active cell is " + activeCell.address);
    });
}).catch(errorHandlerFunction);
```

<span data-ttu-id="52707-116">O método `getSelectedRange()` retorna o intervalo único selecionado atualmente.</span><span class="sxs-lookup"><span data-stu-id="52707-116">The `getSelectedRange()` method returns the currently selected single range.</span></span> <span data-ttu-id="52707-117">Se houver vários intervalos selecionados, será gerado um erro InvalidSelection.</span><span class="sxs-lookup"><span data-stu-id="52707-117">If multiple ranges are selected, an InvalidSelection error is thrown.</span></span> <span data-ttu-id="52707-118">O exemplo a seguir mostra uma chamada para `getSelectedRange()` que, em seguida, define a cor de preenchimento do intervalo como amarelo.</span><span class="sxs-lookup"><span data-stu-id="52707-118">The following example shows a call to `getSelectedRange()` that then sets the range's fill color to yellow.</span></span>

```js
Excel.run(function(context) {
    var range = context.workbook.getSelectedRange();
    range.format.fill.color = "yellow";
    return context.sync();
}).catch(errorHandlerFunction);
```

## <a name="create-a-workbook"></a><span data-ttu-id="52707-119">Criar uma pasta de trabalho</span><span class="sxs-lookup"><span data-stu-id="52707-119">Create a workbook</span></span>

<span data-ttu-id="52707-120">O suplemento pode criar uma nova pasta de trabalho separada da instância do Excel, na qual o suplemento está sendo executado atualmente.</span><span class="sxs-lookup"><span data-stu-id="52707-120">Your add-in can create a new workbook, separate from the Excel instance in which the add-in is currently running.</span></span> <span data-ttu-id="52707-121">O objeto do Excel tem o método `createWorkbook` para esta finalidade.</span><span class="sxs-lookup"><span data-stu-id="52707-121">The Excel object has the `createWorkbook` method for this purpose.</span></span> <span data-ttu-id="52707-122">Quando esse método é chamado, a nova pasta de trabalho é aberta imediatamente e exibida em uma nova instância do Excel.</span><span class="sxs-lookup"><span data-stu-id="52707-122">When this method is called, the new workbook is immediately opened and displayed in a new instance of Excel.</span></span> <span data-ttu-id="52707-123">O suplemento permanece aberto e em execução com a pasta de trabalho anterior.</span><span class="sxs-lookup"><span data-stu-id="52707-123">Your add-in remains open and running with the previous workbook.</span></span>

```js
Excel.createWorkbook();
```

<span data-ttu-id="52707-124">O método `createWorkbook` também cria uma cópia de uma pasta de trabalho existente.</span><span class="sxs-lookup"><span data-stu-id="52707-124">The `createWorkbook` method can also create a copy of an existing workbook.</span></span> <span data-ttu-id="52707-125">O método aceita uma representação de cadeia de caracteres codificada em Base64 de um arquivo .xlsx como parâmetro opcional.</span><span class="sxs-lookup"><span data-stu-id="52707-125">The method accepts a base64-encoded string representation of an .xlsx file as an optional parameter.</span></span> <span data-ttu-id="52707-126">A pasta de trabalho resultante será uma cópia desse arquivo, supondo que o argumento da cadeia de caracteres seja um arquivo. xlsx válido.</span><span class="sxs-lookup"><span data-stu-id="52707-126">The resulting workbook will be a copy of that file, assuming the string argument is a valid .xlsx file.</span></span>

<span data-ttu-id="52707-127">Você pode obter a pasta de trabalho atual do seu complemento como uma cadeia de caracteres codificada com base64 usando o [corte de arquivo](/javascript/api/office/office.document#getfileasync-filetype--options--callback-).</span><span class="sxs-lookup"><span data-stu-id="52707-127">You can get your add-in's current workbook as a base64-encoded string by using [file slicing](/javascript/api/office/office.document#getfileasync-filetype--options--callback-).</span></span> <span data-ttu-id="52707-128">A classe [FileReader](https://developer.mozilla.org/docs/Web/API/FileReader) pode ser usada para converter um arquivo em uma cadeia de caracteres codificada com Base64, como demonstrado no seguinte exemplo.</span><span class="sxs-lookup"><span data-stu-id="52707-128">The [FileReader](https://developer.mozilla.org/docs/Web/API/FileReader) class can be used to convert a file into the required base64-encoded string, as demonstrated in the following example.</span></span>

```js
// Retrieve the external workbook file and set up a `FileReader` object. 
var myFile = document.getElementById("file");
var reader = new FileReader();

reader.onload = (function (event) {
    Excel.run(function (context) {
        // Remove the metadata before the base64-encoded string.
        var startIndex = reader.result.toString().indexOf("base64,");
        var externalWorkbook = reader.result.toString().substr(startIndex + 7);

        Excel.createWorkbook(externalWorkbook);
        return context.sync();
    }).catch(errorHandlerFunction);
});

// Read the file as a data URL so we can parse the base64-encoded string.
reader.readAsDataURL(myFile.files[0]);
```

### <a name="insert-a-copy-of-an-existing-workbook-into-the-current-one"></a><span data-ttu-id="52707-129">Inserir uma cópia de uma pasta de trabalho para a seção atual</span><span class="sxs-lookup"><span data-stu-id="52707-129">Insert a copy of an existing workbook into the current one</span></span>

<span data-ttu-id="52707-130">O exemplo anterior mostra uma nova pasta de trabalho criada a partir de uma pasta de trabalho.</span><span class="sxs-lookup"><span data-stu-id="52707-130">The previous example shows a new workbook being created from an existing workbook.</span></span> <span data-ttu-id="52707-131">Você também pode copiar algumas ou todas de uma pasta de trabalho para a atualmente associada com o suplemento.</span><span class="sxs-lookup"><span data-stu-id="52707-131">You can also copy some or all of an existing workbook into the one currently associated with your add-in.</span></span> <span data-ttu-id="52707-132">Uma [pasta de](/javascript/api/excel/excel.workbook) trabalho tem o método para inserir cópias das planilhas da pasta de trabalho de destino em `insertWorksheetsFromBase64` si.</span><span class="sxs-lookup"><span data-stu-id="52707-132">A [Workbook](/javascript/api/excel/excel.workbook) has the `insertWorksheetsFromBase64` method to insert copies of the target workbook's worksheets into itself.</span></span> <span data-ttu-id="52707-133">O arquivo da outra pasta de trabalho é passado como uma cadeia de caracteres codificada com base64, assim como a `Excel.createWorkbook` chamada.</span><span class="sxs-lookup"><span data-stu-id="52707-133">The other workbook's file is passed as a base64-encoded string, just like the `Excel.createWorkbook` call.</span></span> 

```TypeScript
insertWorksheetsFromBase64(base64File: string, options?: Excel.InsertWorksheetOptions): OfficeExtension.ClientResult<string[]>;
```

> [!IMPORTANT]
> <span data-ttu-id="52707-134">O `insertWorksheetsFromBase64` método é suportado para Excel no Windows, Mac e na Web.</span><span class="sxs-lookup"><span data-stu-id="52707-134">The `insertWorksheetsFromBase64` method is supported for Excel on Windows, Mac, and the web.</span></span> <span data-ttu-id="52707-135">Não há suporte para iOS.</span><span class="sxs-lookup"><span data-stu-id="52707-135">It's not supported for iOS.</span></span> <span data-ttu-id="52707-136">Além disso, Excel na Web este método não dá suporte a planilhas de origem com elementos PivotTable, Chart, Comment ou Slicer.</span><span class="sxs-lookup"><span data-stu-id="52707-136">Additionally, in Excel on the web this method doesn't support source worksheets with PivotTable, Chart, Comment, or Slicer elements.</span></span> <span data-ttu-id="52707-137">Se esses objetos estão presentes, o `insertWorksheetsFromBase64` método retorna o erro em `UnsupportedFeature` Excel na Web.</span><span class="sxs-lookup"><span data-stu-id="52707-137">If those objects are present, the `insertWorksheetsFromBase64` method returns the `UnsupportedFeature` error in Excel on the web.</span></span> 

<span data-ttu-id="52707-138">O exemplo de código a seguir mostra como inserir planilhas de outra pasta de trabalho na pasta de trabalho atual.</span><span class="sxs-lookup"><span data-stu-id="52707-138">The following code sample shows how to insert worksheets from another workbook into the current workbook.</span></span> <span data-ttu-id="52707-139">Este exemplo de código primeiro processa um arquivo de pasta de trabalho com um objeto e extrai uma cadeia de caracteres codificada com base64 e insere essa cadeia de caracteres codificada com base64 na pasta de trabalho [`FileReader`](https://developer.mozilla.org/docs/Web/API/FileReader) atual.</span><span class="sxs-lookup"><span data-stu-id="52707-139">This code sample first processes a workbook file with a [`FileReader`](https://developer.mozilla.org/docs/Web/API/FileReader) object and extracts a base64-encoded string, and then it inserts this base64-encoded string into the current workbook.</span></span> <span data-ttu-id="52707-140">As novas planilhas são inseridas após a planilha chamada **Sheet1**.</span><span class="sxs-lookup"><span data-stu-id="52707-140">The new worksheets are inserted after the worksheet named **Sheet1**.</span></span> <span data-ttu-id="52707-141">Observe que é passado como o parâmetro para a `[]` [propriedade InsertWorksheetOptions.sheetNamesToInsert.](/javascript/api/excel/excel.insertworksheetoptions#sheetNamesToInsert)</span><span class="sxs-lookup"><span data-stu-id="52707-141">Note that `[]` is passed as the parameter for the [InsertWorksheetOptions.sheetNamesToInsert](/javascript/api/excel/excel.insertworksheetoptions#sheetNamesToInsert) property.</span></span> <span data-ttu-id="52707-142">Isso significa que todas as planilhas da pasta de trabalho de destino são inseridas na pasta de trabalho atual.</span><span class="sxs-lookup"><span data-stu-id="52707-142">This means that all the worksheets from the target workbook are inserted into the current workbook.</span></span>

```js
// Retrieve the external workbook file and set up a `FileReader` object. 
var myFile = document.getElementById("file");
var reader = new FileReader();

reader.onload = (event) => {
    Excel.run((context) => {
        // Remove the metadata before the base64-encoded string.
        var startIndex = reader.result.toString().indexOf("base64,");
        var externalWorkbook = reader.result.toString().substr(startIndex + 7);
            
        // Retrieve the current workbook.
        var workbook = context.workbook;
            
        // Set up the insert options. 
        var options = { 
            sheetNamesToInsert: [], // Insert all the worksheets from the source workbook.
            positionType: Excel.WorksheetPositionType.after, // Insert after the `relativeTo` sheet.
            relativeTo: "Sheet1" // The sheet relative to which the other worksheets will be inserted. Used with `positionType`.
        }; 
            
         // Insert the new worksheets into the current workbook.
         workbook.insertWorksheetsFromBase64(externalWorkbook, options);
         return context.sync();
    });
};

// Read the file as a data URL so we can parse the base64-encoded string.
reader.readAsDataURL(myFile.files[0]);
```

## <a name="protect-the-workbooks-structure"></a><span data-ttu-id="52707-143">Protege a estrutura da pasta de trabalho</span><span class="sxs-lookup"><span data-stu-id="52707-143">Protect the workbook's structure</span></span>

<span data-ttu-id="52707-144">O suplemento pode controlar a capacidade de um usuário de editar dados em uma pasta de trabalho.</span><span class="sxs-lookup"><span data-stu-id="52707-144">Your add-in can control a user's ability to edit the workbook's structure.</span></span> <span data-ttu-id="52707-145">A propriedade `protection` do objeto Workbook é um objeto [WorkbookProtection](/javascript/api/excel/excel.workbookprotection) com um método `protect()`.</span><span class="sxs-lookup"><span data-stu-id="52707-145">The Workbook object's `protection` property is a [WorkbookProtection](/javascript/api/excel/excel.workbookprotection) object with a `protect()` method.</span></span> <span data-ttu-id="52707-146">O exemplo a seguir mostra um cenário básico ativando/desativando a proteção da pasta de trabalho.</span><span class="sxs-lookup"><span data-stu-id="52707-146">The following example shows a basic scenario toggling the protection of the workbook's structure.</span></span>

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

<span data-ttu-id="52707-147">O método `protect` aceita um parâmetro opcional de cadeia de caracteres.</span><span class="sxs-lookup"><span data-stu-id="52707-147">The `protect` method accepts an optional string parameter.</span></span> <span data-ttu-id="52707-148">Esta cadeia de caracteres representa a senha necessária para um usuário ignorar a proteção e alterar a estrutura da pasta de trabalho.</span><span class="sxs-lookup"><span data-stu-id="52707-148">This string represents the password needed for a user to bypass protection and change the workbook's structure.</span></span>

<span data-ttu-id="52707-149">A proteção também ser definida no nível da planilha para prevenir a edição de dados indesejada.</span><span class="sxs-lookup"><span data-stu-id="52707-149">Protection can also be set at the worksheet level to prevent unwanted data editing.</span></span> <span data-ttu-id="52707-150">Para saber mais, confira a seção **Proteção de dados** do artigo [Trabalhar com planilhas usando a API JavaScript do Excel](excel-add-ins-worksheets.md#data-protection).</span><span class="sxs-lookup"><span data-stu-id="52707-150">For more information, see the **Data protection** section of the [Work with worksheets using the Excel JavaScript API](excel-add-ins-worksheets.md#data-protection) article.</span></span>

> [!NOTE]
> <span data-ttu-id="52707-151">Para saber mais sobre a proteção de pastas de trabalho no Excel, confira o artigo [Proteger uma pasta de trabalho](https://support.office.com/article/Protect-a-workbook-7E365A4D-3E89-4616-84CA-1931257C1517).</span><span class="sxs-lookup"><span data-stu-id="52707-151">For more information about workbook protection in Excel, see the [Protect a workbook](https://support.office.com/article/Protect-a-workbook-7E365A4D-3E89-4616-84CA-1931257C1517) article.</span></span>

## <a name="access-document-properties"></a><span data-ttu-id="52707-152">Acessar propriedades do documentos</span><span class="sxs-lookup"><span data-stu-id="52707-152">Access document properties</span></span>

<span data-ttu-id="52707-153">Objetos Workbook têm acesso aos metadados dos arquivos do Office, que são conhecidos como [propriedades de documentos](https://support.office.com/article/View-or-change-the-properties-for-an-Office-file-21D604C2-481E-4379-8E54-1DD4622C6B75).</span><span class="sxs-lookup"><span data-stu-id="52707-153">Workbook objects have access to the Office file metadata, which is known as the [document properties](https://support.office.com/article/View-or-change-the-properties-for-an-Office-file-21D604C2-481E-4379-8E54-1DD4622C6B75).</span></span> <span data-ttu-id="52707-154">A propriedade `properties` do objeto Workbook é um objeto [DocumentProperties](/javascript/api/excel/excel.documentproperties) que contém esses valores de metadados.</span><span class="sxs-lookup"><span data-stu-id="52707-154">The Workbook object's `properties` property is a [DocumentProperties](/javascript/api/excel/excel.documentproperties) object containing these metadata values.</span></span> <span data-ttu-id="52707-155">O exemplo a seguir mostra como definir a `author` propriedade.</span><span class="sxs-lookup"><span data-stu-id="52707-155">The following example shows how to set the `author` property.</span></span>

```js
Excel.run(function (context) {
    var docProperties = context.workbook.properties;
    docProperties.author = "Alex";
    return context.sync();
}).catch(errorHandlerFunction);
```

### <a name="custom-properties"></a><span data-ttu-id="52707-156">Propriedades personalizadas</span><span class="sxs-lookup"><span data-stu-id="52707-156">Custom properties</span></span>

<span data-ttu-id="52707-157">Você também pode definir propriedades personalizadas.</span><span class="sxs-lookup"><span data-stu-id="52707-157">You can also define custom properties.</span></span> <span data-ttu-id="52707-158">O objeto DocumentProperties contém uma propriedade `custom` que representa um conjunto de pares de valores-chave para propriedades definidas pelo usuário.</span><span class="sxs-lookup"><span data-stu-id="52707-158">The DocumentProperties object contains a `custom` property that represents a collection of key-value pairs for user-defined properties.</span></span> <span data-ttu-id="52707-159">O exemplo a seguir mostra como criar uma propriedade personalizada chamada **Introduction** com o valor "Olá" e, em seguida, recuperá-la.</span><span class="sxs-lookup"><span data-stu-id="52707-159">The following example shows how to create a custom property named **Introduction** with the value "Hello", then retrieve it.</span></span>

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
    customProperty.load(["key, value"]);

    return context.sync().then(function() {
        console.log("Custom key  : " + customProperty.key); // "Introduction"
        console.log("Custom value : " + customProperty.value); // "Hello"
    });
}).catch(errorHandlerFunction);
```

#### <a name="worksheet-level-custom-properties"></a><span data-ttu-id="52707-160">Propriedades personalizadas no nível da planilha</span><span class="sxs-lookup"><span data-stu-id="52707-160">Worksheet-level custom properties</span></span>

<span data-ttu-id="52707-161">As propriedades personalizadas também podem ser definidas no nível da planilha.</span><span class="sxs-lookup"><span data-stu-id="52707-161">Custom properties can also be set at the worksheet level.</span></span> <span data-ttu-id="52707-162">Elas são semelhantes às propriedades personalizadas no nível do documento, exceto que a mesma chave pode ser repetida em planilhas diferentes.</span><span class="sxs-lookup"><span data-stu-id="52707-162">These are similar to document-level custom properties, except that the same key can be repeated across different worksheets.</span></span> <span data-ttu-id="52707-163">O exemplo a seguir mostra como criar uma propriedade personalizada chamada **WorksheetGroup** com o valor "Alfa" na planilha atual e, em seguida, recuperá-la.</span><span class="sxs-lookup"><span data-stu-id="52707-163">The following example shows how to create a custom property named **WorksheetGroup** with the value "Alpha" on the current worksheet, then retrieve it.</span></span>

```js
Excel.run(function (context) {
    // Add the custom property.
    var customWorksheetProperties = context.workbook.worksheets.getActiveWorksheet().customProperties;
    customWorksheetProperties.add("WorksheetGroup", "Alpha");

    return context.sync();
}).catch(errorHandlerFunction);

[...]

Excel.run(function (context) {
    // Load the keys and values of all custom properties in the current worksheet.
    var worksheet = context.workbook.worksheets.getActiveWorksheet();
    worksheet.load("name");

    var customWorksheetProperties = worksheet.customProperties;
    var customWorksheetProperty = customWorksheetProperties.getItem("WorksheetGroup");
    customWorksheetProperty.load(["key", "value"]);

    return context.sync().then(function() {
        // Log the WorksheetGroup custom property to the console.
        console.log(worksheet.name + ": " + customWorksheetProperty.key); // "WorksheetGroup"
        console.log("  Custom value : " + customWorksheetProperty.value); // "Alpha"
    });
}).catch(errorHandlerFunction);
```

## <a name="access-document-settings"></a><span data-ttu-id="52707-164">Acessar configurações do documentos</span><span class="sxs-lookup"><span data-stu-id="52707-164">Access document settings</span></span>

<span data-ttu-id="52707-165">As configurações da pasta de trabalho são semelhantes ao conjunto de propriedades personalizadas.</span><span class="sxs-lookup"><span data-stu-id="52707-165">A workbook's settings are similar to the collection of custom properties.</span></span> <span data-ttu-id="52707-166">A diferença é que as configurações são exclusivas para um único arquivo do Excel e emparelhamento de suplementos, enquanto que as propriedades estão somente conectadas ao arquivo.</span><span class="sxs-lookup"><span data-stu-id="52707-166">The difference is settings are unique to a single Excel file and add-in pairing, whereas properties are solely connected to the file.</span></span> <span data-ttu-id="52707-167">O exemplo a seguir mostra como criar e acessar uma configuração.</span><span class="sxs-lookup"><span data-stu-id="52707-167">The following example shows how to create and access a setting.</span></span>

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

## <a name="access-application-culture-settings"></a><span data-ttu-id="52707-168">Configurações de cultura de aplicativos do Access</span><span class="sxs-lookup"><span data-stu-id="52707-168">Access application culture settings</span></span>

<span data-ttu-id="52707-169">Uma workbook tem configurações de idioma e cultura que afetam a forma como determinados dados são exibidos.</span><span class="sxs-lookup"><span data-stu-id="52707-169">A workbook has language and culture settings that affect how certain data is displayed.</span></span> <span data-ttu-id="52707-170">Essas configurações podem ajudar a localização de dados quando os usuários do seu complemento estão compartilhando as guias de trabalho em diferentes idiomas e culturas.</span><span class="sxs-lookup"><span data-stu-id="52707-170">These settings can help localize data when your add-in's users are sharing workbooks across different languages and cultures.</span></span> <span data-ttu-id="52707-171">Seu complemento pode usar a análise de cadeia de caracteres para localizar o formato de números, datas e horas com base nas configurações de cultura do sistema para que cada usuário veja dados no formato de sua própria cultura.</span><span class="sxs-lookup"><span data-stu-id="52707-171">Your add-in can use string parsing to localize the format of numbers, dates, and times based on the system culture settings so that each user sees data in their own culture's format.</span></span>

<span data-ttu-id="52707-172">`Application.cultureInfo`define as configurações de cultura do sistema como um [objeto CultureInfo.](/javascript/api/excel/excel.cultureinfo)</span><span class="sxs-lookup"><span data-stu-id="52707-172">`Application.cultureInfo` defines the system culture settings as a [CultureInfo](/javascript/api/excel/excel.cultureinfo) object.</span></span> <span data-ttu-id="52707-173">Isso contém configurações como o separador decimal numérico ou o formato de data.</span><span class="sxs-lookup"><span data-stu-id="52707-173">This contains settings like the numerical decimal separator or the date format.</span></span>

<span data-ttu-id="52707-174">Algumas configurações de cultura podem ser [alteradas por meio da interface do usuário Excel](https://support.office.com/article/Change-the-character-used-to-separate-thousands-or-decimals-c093b545-71cb-4903-b205-aebb9837bd1e)interface do usuário .</span><span class="sxs-lookup"><span data-stu-id="52707-174">Some culture settings can be [changed through the Excel UI](https://support.office.com/article/Change-the-character-used-to-separate-thousands-or-decimals-c093b545-71cb-4903-b205-aebb9837bd1e).</span></span> <span data-ttu-id="52707-175">As configurações do sistema são preservadas no `CultureInfo` objeto.</span><span class="sxs-lookup"><span data-stu-id="52707-175">The system settings are preserved in the `CultureInfo` object.</span></span> <span data-ttu-id="52707-176">Quaisquer alterações locais são mantidas como [propriedades de](/javascript/api/excel/excel.application)nível de aplicativo, como `Application.decimalSeparator` .</span><span class="sxs-lookup"><span data-stu-id="52707-176">Any local changes are kept as [Application](/javascript/api/excel/excel.application)-level properties, such as `Application.decimalSeparator`.</span></span>

<span data-ttu-id="52707-177">O exemplo a seguir altera o caractere separador decimal de uma cadeia numérica de um ',' para o caractere usado pelas configurações do sistema.</span><span class="sxs-lookup"><span data-stu-id="52707-177">The following sample changes the decimal separator character of a numerical string from a ',' to the character used by the system settings.</span></span>

```js
// This will convert a number like "14,37" to "14.37"
// (assuming the system decimal separator is ".").
Excel.run(function (context) {
    var sheet = context.workbook.worksheets.getItem("Sample");
    var decimalSource = sheet.getRange("B2");
    decimalSource.load("values");
    context.application.cultureInfo.numberFormat.load("numberDecimalSeparator");

    return context.sync().then(function() {
        var systemDecimalSeparator =
            context.application.cultureInfo.numberFormat.numberDecimalSeparator;
        var oldDecimalString = decimalSource.values[0][0];

        // This assumes the input column is standardized to use "," as the decimal separator.
        var newDecimalString = oldDecimalString.replace(",", systemDecimalSeparator);

        var resultRange = sheet.getRange("C2");
        resultRange.values = [[newDecimalString]];
        resultRange.format.autofitColumns();
        return context.sync();
    });
});
```

## <a name="add-custom-xml-data-to-the-workbook"></a><span data-ttu-id="52707-178">Adicionar dados XML personalizados à pasta de trabalho</span><span class="sxs-lookup"><span data-stu-id="52707-178">Add custom XML data to the workbook</span></span>

<span data-ttu-id="52707-179">O formato de arquivo Open XML **.xlsx** do Excel permite ao seu suplemento inserir dados XML personalizados na pasta de trabalho.</span><span class="sxs-lookup"><span data-stu-id="52707-179">Excel's Open XML **.xlsx** file format lets your add-in embed custom XML data in the workbook.</span></span> <span data-ttu-id="52707-180">Esses dados persistem na pasta de trabalho, independentemente do suplemento.</span><span class="sxs-lookup"><span data-stu-id="52707-180">This data persists with the workbook, independent of the add-in.</span></span>

<span data-ttu-id="52707-181">Uma pasta de trabalho contém um [CustomXmlPartCollection](/javascript/api/excel/excel.customxmlpartcollection), que é uma lista de [CustomXmlParts](/javascript/api/excel/excel.customxmlpart).</span><span class="sxs-lookup"><span data-stu-id="52707-181">A workbook contains a [CustomXmlPartCollection](/javascript/api/excel/excel.customxmlpartcollection), which is a list of [CustomXmlParts](/javascript/api/excel/excel.customxmlpart).</span></span> <span data-ttu-id="52707-182">Eles oferecem acesso a cadeias de caracteres XML e a uma ID exclusiva correspondente.</span><span class="sxs-lookup"><span data-stu-id="52707-182">These give access to the XML strings and a corresponding unique ID.</span></span> <span data-ttu-id="52707-183">Armazenando essas IDs como configurações, seu suplemento pode manter as teclas para suas partes XML entre sessões.</span><span class="sxs-lookup"><span data-stu-id="52707-183">By storing these IDs as settings, your add-in can maintain the keys to its XML parts between sessions.</span></span>

<span data-ttu-id="52707-184">Os exemplos a seguir mostram como usar partes XML personalizadas.</span><span class="sxs-lookup"><span data-stu-id="52707-184">The following samples show how to use custom XML parts.</span></span> <span data-ttu-id="52707-185">O primeiro bloco de códigos demonstra como inserir dados XML no documento.</span><span class="sxs-lookup"><span data-stu-id="52707-185">The first code block demonstrates how to embed XML data in the document.</span></span> <span data-ttu-id="52707-186">Ele armazena uma lista de revisores e usa as configurações da pasta de trabalho para salvar a `id` do XML para recuperação futura.</span><span class="sxs-lookup"><span data-stu-id="52707-186">It stores a list of reviewers, then uses the workbook's settings to save the XML's `id` for future retrieval.</span></span> <span data-ttu-id="52707-187">O segundo bloco mostra como acessar esse XML mais tarde.</span><span class="sxs-lookup"><span data-stu-id="52707-187">The second block shows how to access that XML later.</span></span> <span data-ttu-id="52707-188">A configuração "ContosoReviewXmlPartId" é carregada e transmitida para `customXmlParts` da pasta de trabalho.</span><span class="sxs-lookup"><span data-stu-id="52707-188">The "ContosoReviewXmlPartId" setting is loaded and passed to the workbook's `customXmlParts`.</span></span> <span data-ttu-id="52707-189">Os dados XML são então impressos no console.</span><span class="sxs-lookup"><span data-stu-id="52707-189">The XML data is then printed to the console.</span></span>

```js
Excel.run(async (context) => {
    // Add reviewer data to the document as XML
    var originalXml = "<Reviewers xmlns='http://schemas.contoso.com/review/1.0'><Reviewer>Juan</Reviewer><Reviewer>Hong</Reviewer><Reviewer>Sally</Reviewer></Reviewers>";
    var customXmlPart = context.workbook.customXmlParts.add(originalXml);
    customXmlPart.load("id");

    return context.sync().then(function() {
        // Store the XML part's ID in a setting
        var settings = context.workbook.settings;
        settings.add("ContosoReviewXmlPartId", customXmlPart.id);
    });
}).catch(errorHandlerFunction);
```

```js
Excel.run(async (context) => {
    // Retrieve the XML part's id from the setting
    var settings = context.workbook.settings;
    var xmlPartIDSetting = settings.getItemOrNullObject("ContosoReviewXmlPartId").load("value");

    return context.sync().then(function () {
        if (xmlPartIDSetting.value) {
            var customXmlPart = context.workbook.customXmlParts.getItem(xmlPartIDSetting.value);
            var xmlBlob = customXmlPart.getXml();

            return context.sync().then(function () {
                // Add spaces to make more human readable in the console
                var readableXML = xmlBlob.value.replace(/></g, "> <");
                console.log(readableXML);
            });
        }
    });
}).catch(errorHandlerFunction);
```

> [!NOTE]
> <span data-ttu-id="52707-190">`CustomXMLPart.namespaceUri` só será preenchido se o elemento XML personalizado de nível superior contiver o atributo `xmlns`.</span><span class="sxs-lookup"><span data-stu-id="52707-190">`CustomXMLPart.namespaceUri` is only populated if the top-level custom XML element contains the `xmlns` attribute.</span></span>

## <a name="control-calculation-behavior"></a><span data-ttu-id="52707-191">Controlar o comportamento do cálculo</span><span class="sxs-lookup"><span data-stu-id="52707-191">Control calculation behavior</span></span>

### <a name="set-calculation-mode"></a><span data-ttu-id="52707-192">Configurar o modo de cálculo</span><span class="sxs-lookup"><span data-stu-id="52707-192">Set calculation mode</span></span>

<span data-ttu-id="52707-193">Por padrão, o Excel recalcula os resultados das fórmulas sempre que uma célula referenciada é alterada.</span><span class="sxs-lookup"><span data-stu-id="52707-193">By default, Excel recalculates formula results whenever a referenced cell is changed.</span></span> <span data-ttu-id="52707-194">O desempenho de seu suplemento pode se beneficiar do ajuste desse comportamento de cálculo.</span><span class="sxs-lookup"><span data-stu-id="52707-194">Your add-in's performance may benefit from adjusting this calculation behavior.</span></span> <span data-ttu-id="52707-195">O objeto Application tem uma propriedade `calculationMode` do tipo `CalculationMode`.</span><span class="sxs-lookup"><span data-stu-id="52707-195">The Application object has a `calculationMode` property of type `CalculationMode`.</span></span> <span data-ttu-id="52707-196">Ele pode ser definido para os seguintes valores.</span><span class="sxs-lookup"><span data-stu-id="52707-196">It can be set to the following values.</span></span>

- <span data-ttu-id="52707-197">`automatic`: O comportamento de recálculo padrão em que o Excel calcula novos resultados das fórmulas sempre que o dado relevante é alterado.</span><span class="sxs-lookup"><span data-stu-id="52707-197">`automatic`: The default recalculation behavior where Excel calculates new formula results every time the relevant data is changed.</span></span>
- <span data-ttu-id="52707-198">`automaticExceptTables`: Igual a `automatic`, exceto que as alterações feitas nos valores em tabelas serão ignoradas.</span><span class="sxs-lookup"><span data-stu-id="52707-198">`automaticExceptTables`: Same as `automatic`, except any changes made to values in tables are ignored.</span></span>
- <span data-ttu-id="52707-199">`manual`: Os cálculos ocorrem somente quando o usuário ou suplemento os solicita.</span><span class="sxs-lookup"><span data-stu-id="52707-199">`manual`: Calculations only occur when the user or add-in requests them.</span></span>

### <a name="set-calculation-type"></a><span data-ttu-id="52707-200">Configurar o tipo de cálculo</span><span class="sxs-lookup"><span data-stu-id="52707-200">Set calculation type</span></span>

<span data-ttu-id="52707-201">O objeto [Application](/javascript/api/excel/excel.application) fornece um método para forçar um recálculo imediato.</span><span class="sxs-lookup"><span data-stu-id="52707-201">The [Application](/javascript/api/excel/excel.application) object provides a method to force an immediate recalculation.</span></span> <span data-ttu-id="52707-202">`Application.calculate(calculationType)` inicia o recálculo manual baseado no `calculationType` especificado.</span><span class="sxs-lookup"><span data-stu-id="52707-202">`Application.calculate(calculationType)` starts a manual recalculation based on the specified `calculationType`.</span></span> <span data-ttu-id="52707-203">Os valores a seguir podem ser especificados.</span><span class="sxs-lookup"><span data-stu-id="52707-203">The following values can be specified.</span></span>

- <span data-ttu-id="52707-204">`full`: Recalcule todas as fórmulas em todas as pastas de trabalho abertas, independentemente de elas terem sido alteradas desde o último recálculo.</span><span class="sxs-lookup"><span data-stu-id="52707-204">`full`: Recalculate all formulas in all open workbooks, regardless of whether they have changed since the last recalculation.</span></span>
- <span data-ttu-id="52707-205">`fullRebuild`: Verifique as fórmulas dependentes e depois recalcule todas as fórmulas em todas as pastas de trabalho abertas, independentemente de elas terem sido alteradas desde o último recálculo.</span><span class="sxs-lookup"><span data-stu-id="52707-205">`fullRebuild`: Check dependent formulas, and then recalculate all formulas in all open workbooks, regardless of whether they have changed since the last recalculation.</span></span>
- <span data-ttu-id="52707-206">`recalculate`: Recalcule as fórmulas que foram alteradas (ou marcadas por programação para recálculo) desde o último cálculo, e as fórmulas dependentes nelas, em todas as pastas de trabalho ativas.</span><span class="sxs-lookup"><span data-stu-id="52707-206">`recalculate`: Recalculate formulas that have changed (or been programmatically marked for recalculation) since the last calculation, and formulas dependent on them, in all active workbooks.</span></span>

> [!NOTE]
> <span data-ttu-id="52707-207">Para saber mais sobre o recálculo, confira o artigo [Alterar o recálculo, a iteração ou a precisão de fórmulas](https://support.office.com/article/change-formula-recalculation-iteration-or-precision-73fc7dac-91cf-4d36-86e8-67124f6bcce4).</span><span class="sxs-lookup"><span data-stu-id="52707-207">For more information about recalculation, see the [Change formula recalculation, iteration, or precision](https://support.office.com/article/change-formula-recalculation-iteration-or-precision-73fc7dac-91cf-4d36-86e8-67124f6bcce4) article.</span></span>

### <a name="temporarily-suspend-calculations"></a><span data-ttu-id="52707-208">Suspender os cálculos temporariamente</span><span class="sxs-lookup"><span data-stu-id="52707-208">Temporarily suspend calculations</span></span>

<span data-ttu-id="52707-209">A API do Excel também permite que os suplementos desativem os cálculos até que `RequestContext.sync()` seja chamado.</span><span class="sxs-lookup"><span data-stu-id="52707-209">The Excel API also lets add-ins turn off calculations until `RequestContext.sync()` is called.</span></span> <span data-ttu-id="52707-210">Isso é feito pelo `suspendApiCalculationUntilNextSync()`.</span><span class="sxs-lookup"><span data-stu-id="52707-210">This is done with `suspendApiCalculationUntilNextSync()`.</span></span> <span data-ttu-id="52707-211">Use esse método quando seu suplemento estiver editando intervalos extensos sem precisar acessar os dados entre as edições.</span><span class="sxs-lookup"><span data-stu-id="52707-211">Use this method when your add-in is editing large ranges without needing to access the data between edits.</span></span>

```js
context.application.suspendApiCalculationUntilNextSync();
```

## <a name="detect-workbook-activation"></a><span data-ttu-id="52707-212">Detectar a ativação de uma agenda de trabalho</span><span class="sxs-lookup"><span data-stu-id="52707-212">Detect workbook activation</span></span>

<span data-ttu-id="52707-213">O seu complemento pode detectar quando uma workbook é ativada.</span><span class="sxs-lookup"><span data-stu-id="52707-213">Your add-in can detect when a workbook is activated.</span></span> <span data-ttu-id="52707-214">Uma workbook fica *inativa* quando o usuário alterna o foco para outra workbook, para outro aplicativo ou (em Excel na Web) para outra guia do navegador da Web.</span><span class="sxs-lookup"><span data-stu-id="52707-214">A workbook becomes *inactive* when the user switches focus to another workbook, to another application, or (in Excel on the web) to another tab of the web browser.</span></span> <span data-ttu-id="52707-215">Uma workbook *é ativada quando* o usuário retorna o foco para a workbook.</span><span class="sxs-lookup"><span data-stu-id="52707-215">A workbook is *activated* when the user returns focus to the workbook.</span></span> <span data-ttu-id="52707-216">A ativação da workbook pode disparar funções de retorno de chamada no seu complemento, como atualizar dados da agenda de trabalho.</span><span class="sxs-lookup"><span data-stu-id="52707-216">The workbook activation can trigger callback functions in your add-in, such as refreshing workbook data.</span></span>

<span data-ttu-id="52707-217">Para detectar quando uma caixa de trabalho é ativada, [registre](excel-add-ins-events.md#register-an-event-handler) um manipulador de eventos para o [evento onActivated](/javascript/api/excel/excel.workbook#onActivated) de uma workbook.</span><span class="sxs-lookup"><span data-stu-id="52707-217">To detect when a workbook is activated, [register an event handler](excel-add-ins-events.md#register-an-event-handler) for the [onActivated](/javascript/api/excel/excel.workbook#onActivated) event of a workbook.</span></span> <span data-ttu-id="52707-218">Os manipuladores de eventos `onActivated` do evento recebem um objeto [WorkbookActivatedEventArgs](/javascript/api/excel/excel.workbookactivatedeventargs) quando o evento é acionado.</span><span class="sxs-lookup"><span data-stu-id="52707-218">Event handlers for the `onActivated` event receive a [WorkbookActivatedEventArgs](/javascript/api/excel/excel.workbookactivatedeventargs) object when the event fires.</span></span>

> [!IMPORTANT]
> <span data-ttu-id="52707-219">O `onActivated` evento não detecta quando uma workbook é aberta.</span><span class="sxs-lookup"><span data-stu-id="52707-219">The `onActivated` event doesn't detect when a workbook is opened.</span></span> <span data-ttu-id="52707-220">Esse evento só detecta quando um usuário alterna o foco de volta para uma workbook já aberta.</span><span class="sxs-lookup"><span data-stu-id="52707-220">This event only detects when a user switches focus back to an already open workbook.</span></span>

<span data-ttu-id="52707-221">O exemplo de código a seguir mostra como registrar o manipulador de eventos `onActivated` e configurar uma função de retorno de chamada.</span><span class="sxs-lookup"><span data-stu-id="52707-221">The following code sample shows how to register the `onActivated` event handler and set up a callback function.</span></span>

```js
Excel.run(function (context) {
    // Retrieve the workbook.
    var workbook = context.workbook;

    // Register the workbook activated event handler.
    workbook.onActivated.add(workbookActivated);

    return context.sync();
});

function workbookActivated(event) {
    Excel.run(function (context) {
        // Retrieve the workbook and load the name.
        var workbook = context.workbook;
        workbook.load("name");
        
        return context.sync().then(function () {
            // Callback function for when the workbook is activated.
            console.log(`The workbook ${workbook.name} was activated.`);
        });
    });
}
```

## <a name="save-the-workbook"></a><span data-ttu-id="52707-222">Salvar a pasta de trabalho</span><span class="sxs-lookup"><span data-stu-id="52707-222">Save the workbook</span></span>

<span data-ttu-id="52707-223">`Workbook.save` salva a pasta de trabalho para armazenamento persistente.</span><span class="sxs-lookup"><span data-stu-id="52707-223">`Workbook.save` saves the workbook to persistent storage.</span></span> <span data-ttu-id="52707-224">O `save` método tem um único parâmetro opcional que pode ser um dos seguintes `saveBehavior` valores.</span><span class="sxs-lookup"><span data-stu-id="52707-224">The `save` method takes a single, optional `saveBehavior` parameter that can be one of the following values.</span></span>

- <span data-ttu-id="52707-225">`Excel.SaveBehavior.save` (padrão): o arquivo será salvo sem solicitar que o usuário especifique o nome do arquivo e local de salvamento.</span><span class="sxs-lookup"><span data-stu-id="52707-225">`Excel.SaveBehavior.save` (default): The file is saved without prompting the user to specify file name and save location.</span></span> <span data-ttu-id="52707-226">Se o arquivo não tiver sido salvo anteriormente, ele será salvo no local padrão.</span><span class="sxs-lookup"><span data-stu-id="52707-226">If the file has not been saved previously, it's saved to the default location.</span></span> <span data-ttu-id="52707-227">Se o arquivo tiver sido salvo anteriormente, ele será salvo no mesmo local.</span><span class="sxs-lookup"><span data-stu-id="52707-227">If the file has been saved previously, it's saved to the same location.</span></span>
- <span data-ttu-id="52707-228">`Excel.SaveBehavior.prompt`: se o arquivo ainda não foi salvo anteriormente, o usuário será solicitado a especificar o nome do arquivo e o local de salvamento.</span><span class="sxs-lookup"><span data-stu-id="52707-228">`Excel.SaveBehavior.prompt`: If file has not been saved previously, the user will be prompted to specify file name and save location.</span></span> <span data-ttu-id="52707-229">Se o arquivo tiver sido salvo anteriormente, ele será salvo no mesmo local sem que o usuário seja solicitado.</span><span class="sxs-lookup"><span data-stu-id="52707-229">If the file has been saved previously, it will be saved to the same location and the user will not be prompted.</span></span>

> [!CAUTION]
> <span data-ttu-id="52707-230">Se o usuário for solicitado a salvar e, em vez disso, cancelar a operação, `save` gera uma exceção.</span><span class="sxs-lookup"><span data-stu-id="52707-230">If the user is prompted to save and cancels the operation, `save` throws an exception.</span></span>

```js
context.workbook.save(Excel.SaveBehavior.prompt);
```

## <a name="close-the-workbook"></a><span data-ttu-id="52707-231">Fechar a pasta de trabalho</span><span class="sxs-lookup"><span data-stu-id="52707-231">Close the workbook</span></span>

<span data-ttu-id="52707-232">`Workbook.close` fecha a pasta de trabalho, além de suplementos que estão associados com a pasta de trabalho (o aplicativo Excel permanece aberto).</span><span class="sxs-lookup"><span data-stu-id="52707-232">`Workbook.close` closes the workbook, along with add-ins that are associated with the workbook (the Excel application remains open).</span></span> <span data-ttu-id="52707-233">O `close` método tem um único parâmetro opcional que pode ser um dos seguintes `closeBehavior` valores.</span><span class="sxs-lookup"><span data-stu-id="52707-233">The `close` method takes a single, optional `closeBehavior` parameter that can be one of the following values.</span></span>

- <span data-ttu-id="52707-234">`Excel.CloseBehavior.save` (padrão): o arquivo será salvo antes de fechar.</span><span class="sxs-lookup"><span data-stu-id="52707-234">`Excel.CloseBehavior.save` (default): The file is saved before closing.</span></span> <span data-ttu-id="52707-235">Se o arquivo não tiver sido salvo anteriormente, o usuário será solicitado a especificar o nome do arquivo e o local para salvá-lo.</span><span class="sxs-lookup"><span data-stu-id="52707-235">If the file has not been saved previously, the user will be prompted to specify file name and save location.</span></span>
- <span data-ttu-id="52707-236">`Excel.CloseBehavior.skipSave`: o arquivo é fechado imediatamente, sem ser salvo.</span><span class="sxs-lookup"><span data-stu-id="52707-236">`Excel.CloseBehavior.skipSave`: The file is immediately closed, without saving.</span></span> <span data-ttu-id="52707-237">Quaisquer alterações não salvas serão perdidas.</span><span class="sxs-lookup"><span data-stu-id="52707-237">Any unsaved changes will be lost.</span></span>

```js
context.workbook.close(Excel.CloseBehavior.save);
```

## <a name="see-also"></a><span data-ttu-id="52707-238">Confira também</span><span class="sxs-lookup"><span data-stu-id="52707-238">See also</span></span>

- [<span data-ttu-id="52707-239">Modelo de objeto JavaScript do Excel em Suplementos do Office</span><span class="sxs-lookup"><span data-stu-id="52707-239">Excel JavaScript object model in Office Add-ins</span></span>](excel-add-ins-core-concepts.md)
- [<span data-ttu-id="52707-240">Trabalhar com planilhas usando a API JavaScript do Excel</span><span class="sxs-lookup"><span data-stu-id="52707-240">Work with worksheets using the Excel JavaScript API</span></span>](excel-add-ins-worksheets.md)
