---
title: Trabalhar com pastas de trabalho usando a API JavaScript do Excel
description: ''
ms.date: 09/26/2019
localization_priority: Priority
ms.openlocfilehash: 66e531a382d467326e5132e60f06c98d414dbb16
ms.sourcegitcommit: 528577145b2cf0a42bc64c56145d661c4d019fb8
ms.translationtype: HT
ms.contentlocale: pt-BR
ms.lasthandoff: 10/02/2019
ms.locfileid: "37353871"
---
# <a name="work-with-workbooks-using-the-excel-javascript-api"></a><span data-ttu-id="f1fc2-102">Trabalhar com pastas de trabalho usando a API JavaScript do Excel</span><span class="sxs-lookup"><span data-stu-id="f1fc2-102">Work with workbooks using the Excel JavaScript API</span></span>

<span data-ttu-id="f1fc2-103">Este artigo fornece exemplos de código que mostram como executar tarefas comuns com pastas de trabalho usando a API JavaScript do Excel.</span><span class="sxs-lookup"><span data-stu-id="f1fc2-103">This article provides code samples that show how to perform common tasks with workbooks using the Excel JavaScript API.</span></span> <span data-ttu-id="f1fc2-104">Para obter a lista completa de propriedades e métodos que o objeto **Workbook** suporta, confira [Objeto Workbook (API JavaScript para Excel)](/javascript/api/excel/excel.workbook).</span><span class="sxs-lookup"><span data-stu-id="f1fc2-104">For the complete list of properties and methods that the **Workbook** object supports, see [Workbook Object (JavaScript API for Excel)](/javascript/api/excel/excel.workbook).</span></span> <span data-ttu-id="f1fc2-105">Este artigo aborda também ações em nível de pasta de trabalho executadas através do objeto [Application](/javascript/api/excel/excel.application).</span><span class="sxs-lookup"><span data-stu-id="f1fc2-105">This article also covers workbook-level actions performed through the [Application](/javascript/api/excel/excel.application) object.</span></span>

<span data-ttu-id="f1fc2-106">O objeto Workbook é o ponto de entrada para que se suplemento interaja com o Excel.</span><span class="sxs-lookup"><span data-stu-id="f1fc2-106">The Workbook object is the entry point for your add-in to interact with Excel.</span></span> <span data-ttu-id="f1fc2-107">Ele mantém conjuntos de planilhas, tabelas, Tabelas Dinâmicas e muito mais, através dos quais os dados do Excel são acessados e alterados.</span><span class="sxs-lookup"><span data-stu-id="f1fc2-107">It maintains collections of worksheets, tables, PivotTables, and more, through which Excel data is accessed and changed.</span></span> <span data-ttu-id="f1fc2-108">O objeto [WorksheetCollection](/javascript/api/excel/excel.worksheetcollection) dá a seu suplemento acesso a todos os dados de pastas de trabalho através de planilhas individuais.</span><span class="sxs-lookup"><span data-stu-id="f1fc2-108">The [WorksheetCollection](/javascript/api/excel/excel.worksheetcollection) object gives your add-in access to all the workbook's data through individual worksheets.</span></span> <span data-ttu-id="f1fc2-109">Especificamente, ele permite seu suplemento adicione planilhas, navegue entre elas e atribua manipuladores a eventos de planilhas.</span><span class="sxs-lookup"><span data-stu-id="f1fc2-109">Specifically, it lets your add-in add worksheets, navigate among them, and assign handlers to worksheet events.</span></span> <span data-ttu-id="f1fc2-110">O artigo [Trabalhar com planilhas usando a API JavaScript do Excel](excel-add-ins-worksheets.md) descreve como acessar e editar planilhas.</span><span class="sxs-lookup"><span data-stu-id="f1fc2-110">The article [Work with worksheets using the Excel JavaScript API](excel-add-ins-worksheets.md) describes how to access and edit worksheets.</span></span>

## <a name="get-the-active-cell-or-selected-range"></a><span data-ttu-id="f1fc2-111">Obter a célula ativa ou o intervalo selecionado</span><span class="sxs-lookup"><span data-stu-id="f1fc2-111">Get the active cell or selected range</span></span>

<span data-ttu-id="f1fc2-112">O objeto Workbook contém dois métodos que obtêm um intervalo de células que o usuário ou o suplemento selecionaram: `getActiveCell()` e `getSelectedRange()`.</span><span class="sxs-lookup"><span data-stu-id="f1fc2-112">The Workbook object contains two methods that get a range of cells the user or add-in has selected: `getActiveCell()` and `getSelectedRange()`.</span></span> <span data-ttu-id="f1fc2-113">`getActiveCell()` obtém a célula ativa da pasta de trabalho como um [objeto Range](/javascript/api/excel/excel.range).</span><span class="sxs-lookup"><span data-stu-id="f1fc2-113">`getActiveCell()` gets the active cell from the workbook as a [Range object](/javascript/api/excel/excel.range).</span></span> <span data-ttu-id="f1fc2-114">O exemplo a seguir mostra uma chamada para `getActiveCell()`, seguida do endereço da célula que está sendo impresso no console.</span><span class="sxs-lookup"><span data-stu-id="f1fc2-114">The following example shows a call to `getActiveCell()`, followed by the cell's address being printed to the console.</span></span>

```js
Excel.run(function (context) {
    var activeCell = context.workbook.getActiveCell();
    activeCell.load("address");

    return context.sync().then(function () {
        console.log("The active cell is " + activeCell.address);
    });
}).catch(errorHandlerFunction);
```

<span data-ttu-id="f1fc2-115">O método `getSelectedRange()` retorna o intervalo único selecionado atualmente.</span><span class="sxs-lookup"><span data-stu-id="f1fc2-115">The `getSelectedRange()` method returns the currently selected single range.</span></span> <span data-ttu-id="f1fc2-116">Se houver vários intervalos selecionados, será gerado um erro InvalidSelection.</span><span class="sxs-lookup"><span data-stu-id="f1fc2-116">If multiple ranges are selected, an InvalidSelection error is thrown.</span></span> <span data-ttu-id="f1fc2-117">O exemplo a seguir mostra uma chamada para `getSelectedRange()` que, em seguida, define a cor de preenchimento do intervalo como amarelo.</span><span class="sxs-lookup"><span data-stu-id="f1fc2-117">The following example shows a call to `getSelectedRange()` that then sets the range's fill color to yellow.</span></span>

```js
Excel.run(function(context) {
    var range = context.workbook.getSelectedRange();
    range.format.fill.color = "yellow";
    return context.sync();
}).catch(errorHandlerFunction);
```

## <a name="create-a-workbook"></a><span data-ttu-id="f1fc2-118">Criar uma pasta de trabalho</span><span class="sxs-lookup"><span data-stu-id="f1fc2-118">Create a workbook</span></span>

<span data-ttu-id="f1fc2-119">O suplemento pode criar uma nova pasta de trabalho separada da instância do Excel, na qual o suplemento está sendo executado atualmente.</span><span class="sxs-lookup"><span data-stu-id="f1fc2-119">Your add-in can create a new workbook, separate from the Excel instance in which the add-in is currently running.</span></span> <span data-ttu-id="f1fc2-120">O objeto do Excel tem o método `createWorkbook` para esta finalidade.</span><span class="sxs-lookup"><span data-stu-id="f1fc2-120">The Excel object has the `createWorkbook` method for this purpose.</span></span> <span data-ttu-id="f1fc2-121">Quando esse método é chamado, a nova pasta de trabalho é aberta imediatamente e exibida em uma nova instância do Excel.</span><span class="sxs-lookup"><span data-stu-id="f1fc2-121">When this method is called, the new workbook is immediately opened and displayed in a new instance of Excel.</span></span> <span data-ttu-id="f1fc2-122">O suplemento permanece aberto e em execução com a pasta de trabalho anterior.</span><span class="sxs-lookup"><span data-stu-id="f1fc2-122">Your add-in remains open and running with the previous workbook.</span></span>

```js
Excel.createWorkbook();
```

<span data-ttu-id="f1fc2-123">O método `createWorkbook` também cria uma cópia de uma pasta de trabalho existente.</span><span class="sxs-lookup"><span data-stu-id="f1fc2-123">The `createWorkbook` method can also create a copy of an existing workbook.</span></span> <span data-ttu-id="f1fc2-124">O método aceita uma representação de cadeia de caracteres codificada em Base64 de um arquivo .xlsx como parâmetro opcional.</span><span class="sxs-lookup"><span data-stu-id="f1fc2-124">The method accepts a base64-encoded string representation of an .xlsx file as an optional parameter.</span></span> <span data-ttu-id="f1fc2-125">A pasta de trabalho resultante será uma cópia desse arquivo, supondo que o argumento da cadeia de caracteres seja um arquivo. xlsx válido.</span><span class="sxs-lookup"><span data-stu-id="f1fc2-125">The resulting workbook will be a copy of that file, assuming the string argument is a valid .xlsx file.</span></span>

<span data-ttu-id="f1fc2-126">Você pode obter a pasta de trabalho atual do suplemento como uma cadeia de caracteres codificada com Base64 usando a [divisão de arquivos](/javascript/api/office/office.document#getfileasync-filetype--options--callback-).</span><span class="sxs-lookup"><span data-stu-id="f1fc2-126">You can get your add-in’s current workbook as a base64-encoded string by using [file slicing](/javascript/api/office/office.document#getfileasync-filetype--options--callback-).</span></span> <span data-ttu-id="f1fc2-127">A classe [FileReader](https://developer.mozilla.org/docs/Web/API/FileReader) pode ser usada para converter um arquivo em uma cadeia de caracteres codificada com Base64, como demonstrado no seguinte exemplo.</span><span class="sxs-lookup"><span data-stu-id="f1fc2-127">The [FileReader](https://developer.mozilla.org/docs/Web/API/FileReader) class can be used to convert a file into the required base64-encoded string, as demonstrated in the following example.</span></span>

```js
var myFile = document.getElementById("file");
var reader = new FileReader();

reader.onload = (function (event) {
    Excel.run(function (context) {
        // strip off the metadata before the base64-encoded string
        var startIndex = reader.result.toString().indexOf("base64,");
        var workbookContents = reader.result.toString().substr(startIndex + 7);

        Excel.createWorkbook(workbookContents);
        return context.sync();
    }).catch(errorHandlerFunction);
});

// read in the file as a data URL so we can parse the base64-encoded string
reader.readAsDataURL(myFile.files[0]);
```

### <a name="insert-a-copy-of-an-existing-workbook-into-the-current-one-preview"></a><span data-ttu-id="f1fc2-128">Inserir uma cópia de uma pasta de trabalho para a seção atual (visualização)</span><span class="sxs-lookup"><span data-stu-id="f1fc2-128">Insert a copy of an existing workbook into the current one</span></span>

> [!NOTE]
> <span data-ttu-id="f1fc2-129">O método`WorksheetCollection.addFromBase64` só está atualmente disponível na versão prévia pública.</span><span class="sxs-lookup"><span data-stu-id="f1fc2-129">The `WorksheetCollection.addFromBase64` method described in this article is currently available only in public preview.</span></span> [!INCLUDE [Information about using preview APIs](../includes/using-excel-preview-apis.md)]

<span data-ttu-id="f1fc2-130">O exemplo anterior mostra uma nova pasta de trabalho criada a partir de uma pasta de trabalho.</span><span class="sxs-lookup"><span data-stu-id="f1fc2-130">The previous example shows a new workbook being created from an existing workbook.</span></span> <span data-ttu-id="f1fc2-131">Você também pode copiar algumas ou todas de uma pasta de trabalho para a atualmente associada com o suplemento.</span><span class="sxs-lookup"><span data-stu-id="f1fc2-131">You can also copy some or all of an existing workbook into the one currently associated with your add-in.</span></span> <span data-ttu-id="f1fc2-132">Uma pasta de trabalho [WorksheetCollection](/javascript/api/excel/excel.worksheetcollection) tem o método `addFromBase64` para inserir cópias de planilhas da pasta de trabalho de destino nela mesma.</span><span class="sxs-lookup"><span data-stu-id="f1fc2-132">A workbook's [WorksheetCollection](/javascript/api/excel/excel.worksheetcollection) has the `addFromBase64` method to insert copies of the target workbook's worksheets into itself.</span></span> <span data-ttu-id="f1fc2-133">O outro arquivo da pasta de trabalho é passado como em cadeia de caracteres codificado em base 64, como a chamada `Excel.createWorkbook`.</span><span class="sxs-lookup"><span data-stu-id="f1fc2-133">The other workbook's file is passed as base64-encoded string, just like the `Excel.createWorkbook` call.</span></span>

```TypeScript
addFromBase64(base64File: string, sheetNamesToInsert?: string[], positionType?: Excel.WorksheetPositionType, relativeTo?: Worksheet | string): OfficeExtension.ClientResult<string[]>;
```

<span data-ttu-id="f1fc2-134">O exemplo a seguir mostra planilhas da pasta de trabalho que estão sendo inseridas em uma pasta de trabalho atual, logo após a planilha ativa.</span><span class="sxs-lookup"><span data-stu-id="f1fc2-134">The following example shows a workbook's worksheets being inserted in the current workbook, directly after the active worksheet.</span></span> <span data-ttu-id="f1fc2-135">Observe que `null` é passado para o parâmetro `sheetNamesToInsert?: string[]`.</span><span class="sxs-lookup"><span data-stu-id="f1fc2-135">Note that `null` is passed for the `sheetNamesToInsert?: string[]` parameter.</span></span> <span data-ttu-id="f1fc2-136">Isso significa que todas as planilhas são inseridas.</span><span class="sxs-lookup"><span data-stu-id="f1fc2-136">This means all the worksheets are being inserted.</span></span>

```js
var myFile = document.getElementById("file");
var reader = new FileReader();

reader.onload = (event) => {
    Excel.run((context) => {
        // strip off the metadata before the base64-encoded string
        var startIndex = reader.result.toString().indexOf("base64,");
        var workbookContents = reader.result.toString().substr(startIndex + 7);

        var sheets = context.workbook.worksheets;
        sheets.addFromBase64(
            workbookContents,
            null, // get all the worksheets
            Excel.WorksheetPositionType.after, // insert them after the worksheet specified by the next parameter
            sheets.getActiveWorksheet() // insert them after the active worksheet
        );
        return context.sync();
    });
};

// read in the file as a data URL so we can parse the base64-encoded string
reader.readAsDataURL(myFile.files[0]);
```

## <a name="protect-the-workbooks-structure"></a><span data-ttu-id="f1fc2-137">Protege a estrutura da pasta de trabalho</span><span class="sxs-lookup"><span data-stu-id="f1fc2-137">Protect the workbook's structure</span></span>

<span data-ttu-id="f1fc2-138">O suplemento pode controlar a capacidade de um usuário de editar dados em uma pasta de trabalho.</span><span class="sxs-lookup"><span data-stu-id="f1fc2-138">Your add-in can control a user's ability to edit the workbook's structure.</span></span> <span data-ttu-id="f1fc2-139">A propriedade `protection` do objeto Workbook é um objeto [WorkbookProtection](/javascript/api/excel/excel.workbookprotection) com um método `protect()`.</span><span class="sxs-lookup"><span data-stu-id="f1fc2-139">The Workbook object's `protection` property is a [WorkbookProtection](/javascript/api/excel/excel.workbookprotection) object with a `protect()` method.</span></span> <span data-ttu-id="f1fc2-140">O exemplo a seguir mostra um cenário básico ativando/desativando a proteção da pasta de trabalho.</span><span class="sxs-lookup"><span data-stu-id="f1fc2-140">The following example shows a basic scenario toggling the protection of the workbook's structure.</span></span>

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

<span data-ttu-id="f1fc2-141">O método `protect` aceita um parâmetro opcional de cadeia de caracteres.</span><span class="sxs-lookup"><span data-stu-id="f1fc2-141">The `protect` method accepts an optional string parameter.</span></span> <span data-ttu-id="f1fc2-142">Esta cadeia de caracteres representa a senha necessária para um usuário ignorar a proteção e alterar a estrutura da pasta de trabalho.</span><span class="sxs-lookup"><span data-stu-id="f1fc2-142">This string represents the password needed for a user to bypass protection and change the workbook's structure.</span></span>

<span data-ttu-id="f1fc2-143">A proteção também ser definida no nível da planilha para prevenir a edição de dados indesejada.</span><span class="sxs-lookup"><span data-stu-id="f1fc2-143">Protection can also be set at the worksheet level to prevent unwanted data editing.</span></span> <span data-ttu-id="f1fc2-144">Para saber mais, confira a seção **Proteção de dados**do artigo [Trabalhar com planilhas usando a API JavaScript do Excel](excel-add-ins-worksheets.md#data-protection).</span><span class="sxs-lookup"><span data-stu-id="f1fc2-144">For more information, see the **Data protection** section of the [Work with worksheets using the Excel JavaScript API](excel-add-ins-worksheets.md#data-protection) article.</span></span>

> [!NOTE]
> <span data-ttu-id="f1fc2-145">Para saber mais sobre a proteção de pastas de trabalho no Excel, confira o artigo [Proteger uma pasta de trabalho](https://support.office.com/article/Protect-a-workbook-7E365A4D-3E89-4616-84CA-1931257C1517).</span><span class="sxs-lookup"><span data-stu-id="f1fc2-145">For more information about workbook protection in Excel, see the [Protect a workbook](https://support.office.com/article/Protect-a-workbook-7E365A4D-3E89-4616-84CA-1931257C1517) article.</span></span>

## <a name="access-document-properties"></a><span data-ttu-id="f1fc2-146">Acessar propriedades do documentos</span><span class="sxs-lookup"><span data-stu-id="f1fc2-146">Access document properties</span></span>

<span data-ttu-id="f1fc2-147">Objetos Workbook têm acesso aos metadados dos arquivos do Office, que são conhecidos como [propriedades de documentos](https://support.office.com/article/View-or-change-the-properties-for-an-Office-file-21D604C2-481E-4379-8E54-1DD4622C6B75).</span><span class="sxs-lookup"><span data-stu-id="f1fc2-147">Workbook objects have access to the Office file metadata, which is known as the [document properties](https://support.office.com/article/View-or-change-the-properties-for-an-Office-file-21D604C2-481E-4379-8E54-1DD4622C6B75).</span></span> <span data-ttu-id="f1fc2-148">A propriedade `properties` do objeto Workbook é um objeto [DocumentProperties](/javascript/api/excel/excel.documentproperties) que contém esses valores de metadados.</span><span class="sxs-lookup"><span data-stu-id="f1fc2-148">The Workbook object's `properties` property is a [DocumentProperties](/javascript/api/excel/excel.documentproperties) object containing these metadata values.</span></span> <span data-ttu-id="f1fc2-149">O exemplo a seguir mostra como definir a propriedade **author**.</span><span class="sxs-lookup"><span data-stu-id="f1fc2-149">The following example shows how to set the **author** property.</span></span>

```js
Excel.run(function (context) {
    var docProperties = context.workbook.properties;
    docProperties.author = "Alex";
    return context.sync();
}).catch(errorHandlerFunction);
```

<span data-ttu-id="f1fc2-150">Você também pode definir propriedades personalizadas.</span><span class="sxs-lookup"><span data-stu-id="f1fc2-150">You can also define custom properties.</span></span> <span data-ttu-id="f1fc2-151">O objeto DocumentProperties contém uma propriedade `custom` que representa um conjunto de pares de valores-chave para propriedades definidas pelo usuário.</span><span class="sxs-lookup"><span data-stu-id="f1fc2-151">The DocumentProperties object contains a `custom` property that represents a collection of key-value pairs for user-defined properties.</span></span> <span data-ttu-id="f1fc2-152">O exemplo a seguir mostra como criar uma propriedade personalizada chamada **Introduction** com o valor "Olá" e, em seguida, recuperá-la.</span><span class="sxs-lookup"><span data-stu-id="f1fc2-152">The following example shows how to create a custom property named **Introduction** with the value "Hello", then retrieve it.</span></span>

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

## <a name="access-document-settings"></a><span data-ttu-id="f1fc2-153">Acessar configurações do documentos</span><span class="sxs-lookup"><span data-stu-id="f1fc2-153">Access document settings</span></span>

<span data-ttu-id="f1fc2-154">As configurações da pasta de trabalho são semelhantes ao conjunto de propriedades personalizadas.</span><span class="sxs-lookup"><span data-stu-id="f1fc2-154">A workbook's settings are similar to the collection of custom properties.</span></span> <span data-ttu-id="f1fc2-155">A diferença é que as configurações são exclusivas para um único arquivo do Excel e emparelhamento de suplementos, enquanto que as propriedades estão somente conectadas ao arquivo.</span><span class="sxs-lookup"><span data-stu-id="f1fc2-155">The difference is settings are unique to a single Excel file and add-in pairing, whereas properties are solely connected to the file.</span></span> <span data-ttu-id="f1fc2-156">O exemplo a seguir mostra como criar e acessar uma configuração.</span><span class="sxs-lookup"><span data-stu-id="f1fc2-156">The following example shows how to create and access a setting.</span></span>

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

## <a name="add-custom-xml-data-to-the-workbook"></a><span data-ttu-id="f1fc2-157">Adicionar dados XML personalizados à pasta de trabalho</span><span class="sxs-lookup"><span data-stu-id="f1fc2-157">Add custom XML data to the workbook</span></span>

<span data-ttu-id="f1fc2-158">O formato de arquivo Open XML **.xlsx** do Excel permite ao seu suplemento inserir dados XML personalizados na pasta de trabalho.</span><span class="sxs-lookup"><span data-stu-id="f1fc2-158">Excel's Open XML **.xlsx** file format lets your add-in embed custom XML data in the workbook.</span></span> <span data-ttu-id="f1fc2-159">Esses dados persistem na pasta de trabalho, independentemente do suplemento.</span><span class="sxs-lookup"><span data-stu-id="f1fc2-159">This data persists with the workbook, independent of the add-in.</span></span>

<span data-ttu-id="f1fc2-160">Uma pasta de trabalho contém um [CustomXmlPartCollection](/javascript/api/excel/excel.customxmlpartcollection), que é uma lista de [CustomXmlParts](/javascript/api/excel/excel.customxmlpart).</span><span class="sxs-lookup"><span data-stu-id="f1fc2-160">A workbook contains a [CustomXmlPartCollection](/javascript/api/excel/excel.customxmlpartcollection), which is a list of [CustomXmlParts](/javascript/api/excel/excel.customxmlpart).</span></span> <span data-ttu-id="f1fc2-161">Eles oferecem acesso a cadeias de caracteres XML e a uma ID exclusiva correspondente.</span><span class="sxs-lookup"><span data-stu-id="f1fc2-161">These give access to the XML strings and a corresponding unique ID.</span></span> <span data-ttu-id="f1fc2-162">Armazenando essas IDs como configurações, seu suplemento pode manter as teclas para suas partes XML entre sessões.</span><span class="sxs-lookup"><span data-stu-id="f1fc2-162">By storing these IDs as settings, your add-in can maintain the keys to its XML parts between sessions.</span></span>

<span data-ttu-id="f1fc2-163">Os exemplos a seguir mostram como usar partes XML personalizadas.</span><span class="sxs-lookup"><span data-stu-id="f1fc2-163">The following samples show how to use custom XML parts.</span></span> <span data-ttu-id="f1fc2-164">O primeiro bloco de códigos demonstra como inserir dados XML no documento.</span><span class="sxs-lookup"><span data-stu-id="f1fc2-164">The first code block demonstrates how to embed XML data in the document.</span></span> <span data-ttu-id="f1fc2-165">Ele armazena uma lista de revisores e usa as configurações da pasta de trabalho para salvar a `id` do XML para recuperação futura.</span><span class="sxs-lookup"><span data-stu-id="f1fc2-165">It stores a list of reviewers, then uses the workbook's settings to save the XML's `id` for future retrieval.</span></span> <span data-ttu-id="f1fc2-166">O segundo bloco mostra como acessar esse XML mais tarde.</span><span class="sxs-lookup"><span data-stu-id="f1fc2-166">The second block shows how to access that XML later.</span></span> <span data-ttu-id="f1fc2-167">A configuração "ContosoReviewXmlPartId" é carregada e transmitida para `customXmlParts` da pasta de trabalho.</span><span class="sxs-lookup"><span data-stu-id="f1fc2-167">The "ContosoReviewXmlPartId" setting is loaded and passed to the workbook's `customXmlParts`.</span></span> <span data-ttu-id="f1fc2-168">Os dados XML são então impressos no console.</span><span class="sxs-lookup"><span data-stu-id="f1fc2-168">The XML data is then printed to the console.</span></span>

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
> <span data-ttu-id="f1fc2-169">`CustomXMLPart.namespaceUri` só será preenchido se o elemento XML personalizado de nível superior contiver o atributo `xmlns`.</span><span class="sxs-lookup"><span data-stu-id="f1fc2-169">`CustomXMLPart.namespaceUri` is only populated if the top-level custom XML element contains the `xmlns` attribute.</span></span>

## <a name="control-calculation-behavior"></a><span data-ttu-id="f1fc2-170">Controlar o comportamento do cálculo</span><span class="sxs-lookup"><span data-stu-id="f1fc2-170">Control calculation behavior</span></span>

### <a name="set-calculation-mode"></a><span data-ttu-id="f1fc2-171">Configurar o modo de cálculo</span><span class="sxs-lookup"><span data-stu-id="f1fc2-171">Set calculation mode</span></span>

<span data-ttu-id="f1fc2-172">Por padrão, o Excel recalcula os resultados das fórmulas sempre que uma célula referenciada é alterada.</span><span class="sxs-lookup"><span data-stu-id="f1fc2-172">By default, Excel recalculates formula results whenever a referenced cell is changed.</span></span> <span data-ttu-id="f1fc2-173">O desempenho de seu suplemento pode se beneficiar do ajuste desse comportamento de cálculo.</span><span class="sxs-lookup"><span data-stu-id="f1fc2-173">Your add-in's performance may benefit from adjusting this calculation behavior.</span></span> <span data-ttu-id="f1fc2-174">O objeto Application tem uma propriedade `calculationMode` do tipo `CalculationMode`.</span><span class="sxs-lookup"><span data-stu-id="f1fc2-174">The Application object has a `calculationMode` property of type `CalculationMode`.</span></span> <span data-ttu-id="f1fc2-175">Esta propriedade pode ser configurada com os seguintes valores:</span><span class="sxs-lookup"><span data-stu-id="f1fc2-175">It can be set to the following values:</span></span>

- <span data-ttu-id="f1fc2-176">`automatic`: O comportamento de recálculo padrão em que o Excel calcula novos resultados das fórmulas sempre que o dado relevante é alterado.</span><span class="sxs-lookup"><span data-stu-id="f1fc2-176">`automatic`: The default recalculation behavior where Excel calculates new formula results every time the relevant data is changed.</span></span>
- <span data-ttu-id="f1fc2-177">`automaticExceptTables`: Igual a `automatic`, exceto que as alterações feitas nos valores em tabelas serão ignoradas.</span><span class="sxs-lookup"><span data-stu-id="f1fc2-177">`automaticExceptTables`: Same as `automatic`, except any changes made to values in tables are ignored.</span></span>
- <span data-ttu-id="f1fc2-178">`manual`: Os cálculos ocorrem somente quando o usuário ou suplemento os solicita.</span><span class="sxs-lookup"><span data-stu-id="f1fc2-178">`manual`: Calculations only occur when the user or add-in requests them.</span></span>

### <a name="set-calculation-type"></a><span data-ttu-id="f1fc2-179">Configurar o tipo de cálculo</span><span class="sxs-lookup"><span data-stu-id="f1fc2-179">Set calculation type</span></span>

<span data-ttu-id="f1fc2-180">O objeto [Application](/javascript/api/excel/excel.application) fornece um método para forçar um recálculo imediato.</span><span class="sxs-lookup"><span data-stu-id="f1fc2-180">The [Application](/javascript/api/excel/excel.application) object provides a method to force an immediate recalculation.</span></span> <span data-ttu-id="f1fc2-181">`Application.calculate(calculationType)` inicia o recálculo manual baseado no `calculationType` especificado.</span><span class="sxs-lookup"><span data-stu-id="f1fc2-181">`Application.calculate(calculationType)` starts a manual recalculation based on the specified `calculationType`.</span></span> <span data-ttu-id="f1fc2-182">Os seguintes valores podem ser especificados:</span><span class="sxs-lookup"><span data-stu-id="f1fc2-182">The following values can be specified:</span></span>

- <span data-ttu-id="f1fc2-183">`full`: Recalcule todas as fórmulas em todas as pastas de trabalho abertas, independentemente de elas terem sido alteradas desde o último recálculo.</span><span class="sxs-lookup"><span data-stu-id="f1fc2-183">`full`: Recalculate all formulas in all open workbooks, regardless of whether they have changed since the last recalculation.</span></span>
- <span data-ttu-id="f1fc2-184">`fullRebuild`: Verifique as fórmulas dependentes e depois recalcule todas as fórmulas em todas as pastas de trabalho abertas, independentemente de elas terem sido alteradas desde o último recálculo.</span><span class="sxs-lookup"><span data-stu-id="f1fc2-184">`fullRebuild`: Check dependent formulas, and then recalculate all formulas in all open workbooks, regardless of whether they have changed since the last recalculation.</span></span>
- <span data-ttu-id="f1fc2-185">`recalculate`: Recalcule as fórmulas que foram alteradas (ou marcadas por programação para recálculo) desde o último cálculo, e as fórmulas dependentes nelas, em todas as pastas de trabalho ativas.</span><span class="sxs-lookup"><span data-stu-id="f1fc2-185">`recalculate`: Recalculate formulas that have changed (or been programmatically marked for recalculation) since the last calculation, and formulas dependent on them, in all active workbooks.</span></span>

> [!NOTE]
> <span data-ttu-id="f1fc2-186">Para saber mais sobre o recálculo, confira o artigo [Alterar o recálculo, a iteração ou a precisão de fórmulas](https://support.office.com/article/change-formula-recalculation-iteration-or-precision-73fc7dac-91cf-4d36-86e8-67124f6bcce4).</span><span class="sxs-lookup"><span data-stu-id="f1fc2-186">For more information about recalculation, see the [Change formula recalculation, iteration, or precision](https://support.office.com/article/change-formula-recalculation-iteration-or-precision-73fc7dac-91cf-4d36-86e8-67124f6bcce4) article.</span></span>

### <a name="temporarily-suspend-calculations"></a><span data-ttu-id="f1fc2-187">Suspender os cálculos temporariamente</span><span class="sxs-lookup"><span data-stu-id="f1fc2-187">Temporarily suspend calculations</span></span>

<span data-ttu-id="f1fc2-188">A API do Excel também permite que os suplementos desativem os cálculos até que `RequestContext.sync()` seja chamado.</span><span class="sxs-lookup"><span data-stu-id="f1fc2-188">The Excel API also lets add-ins turn off calculations until `RequestContext.sync()` is called.</span></span> <span data-ttu-id="f1fc2-189">Isso é feito pelo `suspendApiCalculationUntilNextSync()`.</span><span class="sxs-lookup"><span data-stu-id="f1fc2-189">This is done with `suspendApiCalculationUntilNextSync()`.</span></span> <span data-ttu-id="f1fc2-190">Use esse método quando seu suplemento estiver editando intervalos extensos sem precisar acessar os dados entre as edições.</span><span class="sxs-lookup"><span data-stu-id="f1fc2-190">Use this method when your add-in is editing large ranges without needing to access the data between edits.</span></span>

```js
context.application.suspendApiCalculationUntilNextSync();
```

## <a name="comments-preview"></a><span data-ttu-id="f1fc2-191">Comentários (visualização)</span><span class="sxs-lookup"><span data-stu-id="f1fc2-191">Comments (preview)</span></span>

> [!NOTE]
> <span data-ttu-id="f1fc2-192">As APIs de comentário estão disponíveis atualmente apenas na visualização pública.</span><span class="sxs-lookup"><span data-stu-id="f1fc2-192">The following events are currently available only in public preview.</span></span> [!INCLUDE [Information about using preview APIs](../includes/using-excel-preview-apis.md)]

<span data-ttu-id="f1fc2-193">Todos os [comentários](https://support.office.com/article/insert-comments-and-notes-in-excel-bdcc9f5d-38e2-45b4-9a92-0b2b5c7bf6f8) em uma pasta de trabalho são acompanhados pela propriedade `Workbook.comments`.</span><span class="sxs-lookup"><span data-stu-id="f1fc2-193">All [comments](https://support.office.com/article/insert-comments-and-notes-in-excel-bdcc9f5d-38e2-45b4-9a92-0b2b5c7bf6f8) within a workbook are tracked by the `Workbook.comments` property.</span></span> <span data-ttu-id="f1fc2-194">Isso inclui comentários criados por usuários e comentários criados por seu suplemento.</span><span class="sxs-lookup"><span data-stu-id="f1fc2-194">This includes comments created by users and also comments created by your add-in.</span></span> <span data-ttu-id="f1fc2-195">A propriedade `Workbook.comments` é um objeto [CommentCollection](/javascript/api/excel/excel.commentcollection) que contém um conjunto de objetos [Comentário](/javascript/api/excel/excel.comment).</span><span class="sxs-lookup"><span data-stu-id="f1fc2-195">The `Workbook.comments` property is a [CommentCollection](/javascript/api/excel/excel.commentcollection) object that contains a collection of [Comment](/javascript/api/excel/excel.comment) objects.</span></span>

<span data-ttu-id="f1fc2-196">Para adicionar comentários a uma pasta de trabalho, use o método `CommentCollection.add`, passando a célula onde o comentário será adicionado, como uma cadeia de caracteres ou um objeto [Range](/javascript/api/excel/excel.range) e o texto do comentário, como uma cadeia de caracteres.</span><span class="sxs-lookup"><span data-stu-id="f1fc2-196">To add comments to a workbook, use the `CommentCollection.add` method, passing in the cell where the comment will be added, as either a string or [Range](/javascript/api/excel/excel.range) object, and the comment's text, as a string.</span></span> <span data-ttu-id="f1fc2-197">O exemplo a seguir adiciona um comentário à célula **A2**.</span><span class="sxs-lookup"><span data-stu-id="f1fc2-197">The following code sample adds a comment to cell **A2**.</span></span>

```js
Excel.run(function (context) {
    var comments = context.workbook.comments;

    // Note that an InvalidArgument error will be thrown if multiple cells passed to `Comment.add`.
    comments.add("A2", "TODO: add data.");
    return context.sync();
});
```

<span data-ttu-id="f1fc2-198">Cada comentário contém metadados sobre a criação, como o autor e a data de criação.</span><span class="sxs-lookup"><span data-stu-id="f1fc2-198">Each comment contains metadata about its creation, such as the author and creation date.</span></span> <span data-ttu-id="f1fc2-199">Os comentários criados por seu suplemento são considerados criados pelo usuário atual.</span><span class="sxs-lookup"><span data-stu-id="f1fc2-199">Comments created by your add-in are considered to be authored by the current user.</span></span> <span data-ttu-id="f1fc2-200">O exemplo a seguir mostra como exibir o email do autor, o nome do autor e a data de criação de um comentário em **A2**.</span><span class="sxs-lookup"><span data-stu-id="f1fc2-200">The following sample shows how to display the author's email, author's name, and creation date of a comment at **A2**.</span></span>

```js
Excel.run(function (context) {
    // Get the comment at cell A2.
    var comment = context.workbook.comments.getItemByCell("Comments!A2");
    comment.load(["authorEmail", "authorName", "creationDate"]);
    return context.sync().then(function () {
        console.log(`${comment.creationDate.toDateString()}: ${comment.authorName} (${comment.authorEmail})`);
    });
});
```

<span data-ttu-id="f1fc2-201">Cada comentário contém zero ou mais respostas.</span><span class="sxs-lookup"><span data-stu-id="f1fc2-201">Each comment contains zero or more replies.</span></span> <span data-ttu-id="f1fc2-202">os objetos `Comment` têm uma propriedade `replies`, que é [CommentReplyCollection](/javascript/api/excel/excel.commentreplycollection) que contém objetos [CommentReply](/javascript/api/excel/excel.commentreply).</span><span class="sxs-lookup"><span data-stu-id="f1fc2-202">`Comment` objects have a `replies` property, which is a [CommentReplyCollection](/javascript/api/excel/excel.commentreplycollection) that contains [CommentReply](/javascript/api/excel/excel.commentreply) objects.</span></span> <span data-ttu-id="f1fc2-203">Para adicionar uma resposta a um comentário, use o método `CommentReplyCollection.add`, passando o texto da resposta.</span><span class="sxs-lookup"><span data-stu-id="f1fc2-203">To add a reply to a comment, use the `CommentReplyCollection.add` method, passing in the text of the reply.</span></span> <span data-ttu-id="f1fc2-204">As respostas são exibidas na ordem em que são adicionadas.</span><span class="sxs-lookup"><span data-stu-id="f1fc2-204">Replies are displayed in the order they are added.</span></span> <span data-ttu-id="f1fc2-205">O exemplo a seguir adiciona uma resposta ao primeiro comentário da pasta de trabalho.</span><span class="sxs-lookup"><span data-stu-id="f1fc2-205">The following code sample adds a data series to the first chart in the worksheet.</span></span>

```js
Excel.run(function (context) {
    // Get the first comment added to the workbook.
    var comment = context.workbook.comments.getItemAt(0);
    comment.replies.add("Thanks for the reminder!");
    return context.sync();
});
```

<span data-ttu-id="f1fc2-206">Para editar um comentário ou uma resposta de comentário, defina uma propriedade`Comment.content` e uma propriedade`CommentReply.content`.</span><span class="sxs-lookup"><span data-stu-id="f1fc2-206">To edit a comment or comment reply, set its `Comment.content` property or `CommentReply.content` property.</span></span> <span data-ttu-id="f1fc2-207">Para excluir um comentário ou uma resposta de comentário, use o método `Comment.delete` ou o método`CommentReply.delete`.</span><span class="sxs-lookup"><span data-stu-id="f1fc2-207">To delete a comment or comment reply, use the `Comment.delete` method or `CommentReply.delete` method.</span></span> <span data-ttu-id="f1fc2-208">Excluir um comentário também exclui todas as respostas associadas a esse comentário.</span><span class="sxs-lookup"><span data-stu-id="f1fc2-208">Deleting a comment also deletes all the replies associated with that comment.</span></span>

> [!TIP]
> <span data-ttu-id="f1fc2-209">Os comentários também podem ser gerenciados no nível da [planilha](/javascript/api/excel/excel.worksheet) usando as mesmas técnicas.</span><span class="sxs-lookup"><span data-stu-id="f1fc2-209">Comments can also be managed at the [Worksheet](/javascript/api/excel/excel.worksheet) level using the same techniques.</span></span>

## <a name="save-the-workbook-preview"></a><span data-ttu-id="f1fc2-210">Salvar a pasta de trabalho (visualização)</span><span class="sxs-lookup"><span data-stu-id="f1fc2-210">Save the workbook</span></span>

> [!NOTE]
> <span data-ttu-id="f1fc2-211">O método`Workbook.save` só está atualmente disponível na versão prévia pública.</span><span class="sxs-lookup"><span data-stu-id="f1fc2-211">The `Workbook.save` method described in this article is currently available only in public preview.</span></span> [!INCLUDE [Information about using preview APIs](../includes/using-excel-preview-apis.md)]

<span data-ttu-id="f1fc2-212">`Workbook.save` salva a pasta de trabalho para armazenamento persistente.</span><span class="sxs-lookup"><span data-stu-id="f1fc2-212">`Workbook.save` saves the workbook to persistent storage .</span></span> <span data-ttu-id="f1fc2-213">O método `save` usa um parâmetro simples e opcional `saveBehavior` que pode ter um dos seguintes valores:</span><span class="sxs-lookup"><span data-stu-id="f1fc2-213">The `save` method takes a single, optional parameter that can be one of the following values:</span></span>

- <span data-ttu-id="f1fc2-214">`Excel.SaveBehavior.save` (padrão): o arquivo será salvo sem solicitar que o usuário especifique o nome do arquivo e local de salvamento.</span><span class="sxs-lookup"><span data-stu-id="f1fc2-214">`Excel.SaveBehavior.save` (default): The file is saved without prompting the user to specify file name and save location.</span></span> <span data-ttu-id="f1fc2-215">Se o arquivo não tiver sido salvo anteriormente, ele será salvo no local padrão.</span><span class="sxs-lookup"><span data-stu-id="f1fc2-215">If the file has not been saved previously, it's saved to the default location.</span></span> <span data-ttu-id="f1fc2-216">Se o arquivo tiver sido salvo anteriormente, ele será salvo no mesmo local.</span><span class="sxs-lookup"><span data-stu-id="f1fc2-216">If the file has been saved previously, it's saved to the same location.</span></span>
- <span data-ttu-id="f1fc2-217">`Excel.SaveBehavior.prompt`: se o arquivo ainda não foi salvo anteriormente, o usuário será solicitado a especificar o nome do arquivo e o local de salvamento.</span><span class="sxs-lookup"><span data-stu-id="f1fc2-217">`Excel.SaveBehavior.prompt`: If file has not been saved previously, the user will be prompted to specify file name and save location.</span></span> <span data-ttu-id="f1fc2-218">Se o arquivo tiver sido salvo anteriormente, ele será salvo no mesmo local sem que o usuário seja solicitado.</span><span class="sxs-lookup"><span data-stu-id="f1fc2-218">If the file has been saved previously, it will be saved to the same location and the user will not be prompted.</span></span>

> [!CAUTION]
> <span data-ttu-id="f1fc2-219">Se o usuário for solicitado a salvar e, em vez disso, cancelar a operação, `save` gera uma exceção.</span><span class="sxs-lookup"><span data-stu-id="f1fc2-219">If the user is prompted to save and cancels the operation, `save` throws an exception.</span></span>

```js
context.workbook.save(Excel.SaveBehavior.prompt);
```

## <a name="close-the-workbook-preview"></a><span data-ttu-id="f1fc2-220">Feche a pasta de trabalho (visualização)</span><span class="sxs-lookup"><span data-stu-id="f1fc2-220">Close the workbook</span></span>

> [!NOTE]
> <span data-ttu-id="f1fc2-221">O método`Workbook.close` só está atualmente disponível na versão prévia pública.</span><span class="sxs-lookup"><span data-stu-id="f1fc2-221">The `Workbook.close` method described in this article is currently available only in public preview.</span></span> [!INCLUDE [Information about using preview APIs](../includes/using-excel-preview-apis.md)]

<span data-ttu-id="f1fc2-222">`Workbook.close` fecha a pasta de trabalho, além de suplementos que estão associados com a pasta de trabalho (o aplicativo Excel permanece aberto).</span><span class="sxs-lookup"><span data-stu-id="f1fc2-222">`Workbook.close` closes the workbook, along with add-ins that are associated with the workbook (the Excel application remains open).</span></span> <span data-ttu-id="f1fc2-223">O método `close` usa um parâmetro simples e opcional `closeBehavior` que pode ter um dos seguintes valores:</span><span class="sxs-lookup"><span data-stu-id="f1fc2-223">The `close` method takes a single, optional parameter that can be one of the following values:</span></span>

- <span data-ttu-id="f1fc2-224">`Excel.CloseBehavior.save` (padrão): o arquivo será salvo antes de fechar.</span><span class="sxs-lookup"><span data-stu-id="f1fc2-224">`Excel.CloseBehavior.save` (default): The file is saved before closing.</span></span> <span data-ttu-id="f1fc2-225">Se o arquivo não tiver sido salvo anteriormente, o usuário será solicitado a especificar o nome do arquivo e o local para salvá-lo.</span><span class="sxs-lookup"><span data-stu-id="f1fc2-225">If the file has not been saved previously, the user will be prompted to specify file name and save location.</span></span>
- <span data-ttu-id="f1fc2-226">`Excel.CloseBehavior.skipSave`: o arquivo é fechado imediatamente, sem ser salvo.</span><span class="sxs-lookup"><span data-stu-id="f1fc2-226">`Excel.CloseBehavior.skipSave`: The file is immediately closed, without saving.</span></span> <span data-ttu-id="f1fc2-227">Quaisquer alterações não salvas serão perdidas.</span><span class="sxs-lookup"><span data-stu-id="f1fc2-227">Any unsaved changes will be lost.</span></span>

```js
context.workbook.close(Excel.CloseBehavior.save);
```

## <a name="see-also"></a><span data-ttu-id="f1fc2-228">Confira também</span><span class="sxs-lookup"><span data-stu-id="f1fc2-228">See also</span></span>

- [<span data-ttu-id="f1fc2-229">Conceitos fundamentais de programação com a API JavaScript do Excel</span><span class="sxs-lookup"><span data-stu-id="f1fc2-229">Fundamental programming concepts with the Excel JavaScript API</span></span>](excel-add-ins-core-concepts.md)
- [<span data-ttu-id="f1fc2-230">Trabalhar com planilhas usando a API JavaScript do Excel</span><span class="sxs-lookup"><span data-stu-id="f1fc2-230">Work with worksheets using the Excel JavaScript API</span></span>](excel-add-ins-worksheets.md)
- [<span data-ttu-id="f1fc2-231">Trabalhar com intervalos usando a API JavaScript do Excel</span><span class="sxs-lookup"><span data-stu-id="f1fc2-231">Work with ranges using the Excel JavaScript API</span></span>](excel-add-ins-ranges.md)
