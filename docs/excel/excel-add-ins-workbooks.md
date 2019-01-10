---
title: Trabalhar com pastas de trabalho usando a API JavaScript do Excel
description: ''
ms.date: 1/7/2019
ms.openlocfilehash: db32cf0c847d578fb909d9ad97a3a75ef3f97eee
ms.sourcegitcommit: 9afcb1bb295ec0c8940ed3a8364dbac08ef6b382
ms.translationtype: HT
ms.contentlocale: pt-BR
ms.lasthandoff: 01/08/2019
ms.locfileid: "27770585"
---
# <a name="work-with-workbooks-using-the-excel-javascript-api"></a><span data-ttu-id="05a2b-102">Trabalhar com pastas de trabalho usando a API JavaScript do Excel</span><span class="sxs-lookup"><span data-stu-id="05a2b-102">Work with workbooks using the Excel JavaScript API</span></span>

<span data-ttu-id="05a2b-103">Este artigo fornece exemplos de código que mostram como executar tarefas comuns com pastas de trabalho usando a API JavaScript do Excel.</span><span class="sxs-lookup"><span data-stu-id="05a2b-103">This article provides code samples that show how to perform common tasks with workbooks using the Excel JavaScript API.</span></span> <span data-ttu-id="05a2b-104">Para obter a lista completa de propriedades e métodos que o objeto **Workbook** suporta, confira [Objeto Workbook (API JavaScript para Excel)](/javascript/api/excel/excel.workbook).</span><span class="sxs-lookup"><span data-stu-id="05a2b-104">For the complete list of properties and methods that the **Workbook** object supports, see [Workbook Object (JavaScript API for Excel)](/javascript/api/excel/excel.workbook).</span></span> <span data-ttu-id="05a2b-105">Este artigo aborda também ações em nível de pasta de trabalho executadas através do objeto [Application](/javascript/api/excel/excel.application).</span><span class="sxs-lookup"><span data-stu-id="05a2b-105">This article also covers workbook-level actions performed through the [Application](/javascript/api/excel/excel.application) object.</span></span>

<span data-ttu-id="05a2b-106">O objeto Workbook é o ponto de entrada para que se suplemento interaja com o Excel.</span><span class="sxs-lookup"><span data-stu-id="05a2b-106">The Workbook object is the entry point for your add-in to interact with Excel.</span></span> <span data-ttu-id="05a2b-107">Ele mantém conjuntos de planilhas, tabelas, Tabelas Dinâmicas e muito mais, através dos quais os dados do Excel são acessados e alterados.</span><span class="sxs-lookup"><span data-stu-id="05a2b-107">It maintains collections of worksheets, tables, PivotTables, and more, through which Excel data is accessed and changed.</span></span> <span data-ttu-id="05a2b-108">O objeto [WorksheetCollection](/javascript/api/excel/excel.worksheetcollection) dá a seu suplemento acesso a todos os dados de pastas de trabalho através de planilhas individuais.</span><span class="sxs-lookup"><span data-stu-id="05a2b-108">The [WorksheetCollection](/javascript/api/excel/excel.worksheetcollection) object gives your add-in access to all the workbook's data through indivual worksheets.</span></span> <span data-ttu-id="05a2b-109">Especificamente, ele permite seu suplemento adicione planilhas, navegue entre elas e atribua manipuladores a eventos de planilhas.</span><span class="sxs-lookup"><span data-stu-id="05a2b-109">Specifically, it lets your add-in add worksheets, navigate among them, and assign handlers to worksheet events.</span></span> <span data-ttu-id="05a2b-110">O artigo [Trabalhar com planilhas usando a API JavaScript do Excel](excel-add-ins-worksheets.md) descreve como acessar e editar planilhas.</span><span class="sxs-lookup"><span data-stu-id="05a2b-110">The article [Work with worksheets using the Excel JavaScript API](excel-add-ins-worksheets.md) describes how to access and edit worksheets.</span></span>

## <a name="get-the-active-cell-or-selected-range"></a><span data-ttu-id="05a2b-111">Obter a célula ativa ou o intervalo selecionado</span><span class="sxs-lookup"><span data-stu-id="05a2b-111">Get the active cell or selected range</span></span>

<span data-ttu-id="05a2b-112">O objeto Workbook contém dois métodos que obtêm um intervalo de células que o usuário ou o suplemento selecionaram: `getActiveCell()` e `getSelectedRange()`.</span><span class="sxs-lookup"><span data-stu-id="05a2b-112">The Workbook object contains two methods that get a range of cells the user or add-in has selected: `getActiveCell()` and `getSelectedRange()`.</span></span> <span data-ttu-id="05a2b-113">`getActiveCell()` obtém a célula ativa da pasta de trabalho como um [objeto Range](/javascript/api/excel/excel.range).</span><span class="sxs-lookup"><span data-stu-id="05a2b-113">`getActiveCell()` gets the active cell from the workbook as a [Range object](/javascript/api/excel/excel.range).</span></span> <span data-ttu-id="05a2b-114">O exemplo a seguir mostra uma chamada para `getActiveCell()`, seguida do endereço da célula que está sendo impresso no console.</span><span class="sxs-lookup"><span data-stu-id="05a2b-114">The following example shows a call to `getActiveCell()`, followed by the cell's address being printed to the console.</span></span>

```js
Excel.run(function (context) {
    var activeCell = context.workbook.getActiveCell();
    activeCell.load("address");

    return context.sync().then(function () {
        console.log("The active cell is " + activeCell.address);
    });
}).catch(errorHandlerFunction);
```

<span data-ttu-id="05a2b-115">O método `getSelectedRange()` retorna o intervalo único selecionado atualmente.</span><span class="sxs-lookup"><span data-stu-id="05a2b-115">The `getSelectedRange()` method returns the currently selected single range.</span></span> <span data-ttu-id="05a2b-116">Se houver vários intervalos selecionados, será gerado um erro InvalidSelection.</span><span class="sxs-lookup"><span data-stu-id="05a2b-116">If multiple ranges are selected, an InvalidSelection error is thrown.</span></span> <span data-ttu-id="05a2b-117">O exemplo a seguir mostra uma chamada para `getSelectedRange()` que, em seguida, define a cor de preenchimento do intervalo como amarelo.</span><span class="sxs-lookup"><span data-stu-id="05a2b-117">The following example shows a call to `getSelectedRange()` that then sets the range's fill color to yellow.</span></span>

```js
Excel.run(function(context) {
    var range = context.workbook.getSelectedRange();
    range.format.fill.color = "yellow";
    return context.sync();
}).catch(errorHandlerFunction);
```

## <a name="create-a-workbook"></a><span data-ttu-id="05a2b-118">Criar uma pasta de trabalho</span><span class="sxs-lookup"><span data-stu-id="05a2b-118">Create a workbook</span></span>

<span data-ttu-id="05a2b-119">O suplemento pode criar uma nova pasta de trabalho separada da instância do Excel, na qual o suplemento está sendo executado atualmente.</span><span class="sxs-lookup"><span data-stu-id="05a2b-119">Your add-in can create a new workbook, separate from the Excel instance in which the add-in is currently running.</span></span> <span data-ttu-id="05a2b-120">O objeto do Excel tem o método `createWorkbook` para esta finalidade.</span><span class="sxs-lookup"><span data-stu-id="05a2b-120">The Excel object has the `createWorkbook` method for this purpose.</span></span> <span data-ttu-id="05a2b-121">Quando esse método é chamado, a nova pasta de trabalho é aberta imediatamente e exibida em uma nova instância do Excel.</span><span class="sxs-lookup"><span data-stu-id="05a2b-121">When this method is called, the new workbook is immediately opened and displayed in a new instance of Excel.</span></span> <span data-ttu-id="05a2b-122">O suplemento permanece aberto e em execução com a pasta de trabalho anterior.</span><span class="sxs-lookup"><span data-stu-id="05a2b-122">Your add-in remains open and running with the previous workbook.</span></span>

```js
Excel.createWorkbook();
```

<span data-ttu-id="05a2b-123">O método `createWorkbook` também cria uma cópia de uma pasta de trabalho existente.</span><span class="sxs-lookup"><span data-stu-id="05a2b-123">The `createWorkbook` method can also create a copy of an existing workbook.</span></span> <span data-ttu-id="05a2b-124">O método aceita uma representação de cadeia de caracteres codificada em Base64 de um arquivo .xlsx como parâmetro opcional.</span><span class="sxs-lookup"><span data-stu-id="05a2b-124">The method accepts a base64-encoded string representation of an .xlsx file as an optional parameter.</span></span> <span data-ttu-id="05a2b-125">A pasta de trabalho resultante será uma cópia desse arquivo, supondo que o argumento da cadeia de caracteres seja um arquivo. xlsx válido.</span><span class="sxs-lookup"><span data-stu-id="05a2b-125">The resulting workbook will be a copy of that file, assuming the string argument is a valid .xlsx file.</span></span>

<span data-ttu-id="05a2b-126">Você pode obter a pasta de trabalho atual do suplemento como uma cadeia de caracteres codificada com Base64 usando a [divisão de arquivos](/javascript/api/office/office.document#getfileasync-filetype--options--callback-).</span><span class="sxs-lookup"><span data-stu-id="05a2b-126">You can get your add-in’s current workbook as a base64-encoded string by using [file slicing](/javascript/api/office/office.document#getfileasync-filetype--options--callback-).</span></span> <span data-ttu-id="05a2b-127">A classe [FileReader](https://developer.mozilla.org/docs/Web/API/FileReader) pode ser usada para converter um arquivo em uma cadeia de caracteres codificada com Base64, como demonstrado no seguinte exemplo.</span><span class="sxs-lookup"><span data-stu-id="05a2b-127">The [FileReader](https://developer.mozilla.org/docs/Web/API/FileReader) class can be used to convert a file into the required base64-encoded string, as demonstrated in the following example.</span></span>

```js
var myFile = document.getElementById("file");
var reader = new FileReader();

reader.onload = (function (event) {
    Excel.run(function (context) {
        // strip off the metadata before the base64-encoded string
        var startIndex = event.target.result.indexOf("base64,");
        var workbookContents = event.target.result.substr(startIndex + 7);

        Excel.createWorkbook(workbookContents);
        return context.sync();
    }).catch(errorHandlerFunction);
});

// read in the file as a data URL so we can parse the base64-encoded string
reader.readAsDataURL(myFile.files[0]);
```

### <a name="insert-a-copy-of-an-existing-workbook-into-the-current-one"></a><span data-ttu-id="05a2b-128">Inserir uma cópia de uma pasta de trabalho para a seção atual</span><span class="sxs-lookup"><span data-stu-id="05a2b-128">Insert a copy of an existing workbook into the current one</span></span>

> [!NOTE]
> <span data-ttu-id="05a2b-129">A função `WorksheetCollection.addFromBase64` só está disponível atualmente na versão visualização pública (beta).</span><span class="sxs-lookup"><span data-stu-id="05a2b-129">The `WorksheetCollection.addFromBase64` function is currently available only in public preview (beta).</span></span> <span data-ttu-id="05a2b-130">Para usar esse recurso, você deve usar a biblioteca beta do CDN do Office.js: https://appsforoffice.microsoft.com/lib/beta/hosted/office.js.</span><span class="sxs-lookup"><span data-stu-id="05a2b-130">To use this feature, you must use the beta library of the Office.js CDN: https://appsforoffice.microsoft.com/lib/beta/hosted/office.js.</span></span>
> <span data-ttu-id="05a2b-131">Se você estiver usando o TypeScript ou se seu editor de código usar arquivos de definição de tipo do TypeScript do IntelliSense, use https://appsforoffice.microsoft.com/lib/beta/hosted/office.d.ts.</span><span class="sxs-lookup"><span data-stu-id="05a2b-131">If you are using TypeScript or your code editor uses TypeScript type definition files for IntelliSense, use https://appsforoffice.microsoft.com/lib/beta/hosted/office.d.ts.</span></span>

<span data-ttu-id="05a2b-132">O exemplo anterior mostra uma nova pasta de trabalho criada a partir de uma pasta de trabalho.</span><span class="sxs-lookup"><span data-stu-id="05a2b-132">The previous example shows a new workbook being created from an existing workbook.</span></span> <span data-ttu-id="05a2b-133">Você também pode copiar algumas ou todas de uma pasta de trabalho para a atualmente associada com o suplemento.</span><span class="sxs-lookup"><span data-stu-id="05a2b-133">You can also copy some or all of an existing workbook into the one currently associated with your add-in.</span></span> <span data-ttu-id="05a2b-134">Uma pasta de trabalho [WorksheetCollection](/javascript/api/excel/excel.worksheetcollection) tem o método `addFromBase64` para inserir cópias de planilhas da pasta de trabalho de destino nela mesma.</span><span class="sxs-lookup"><span data-stu-id="05a2b-134">A workbook's [WorksheetCollection](/javascript/api/excel/excel.worksheetcollection) has the `addFromBase64` method to insert copies of the target workbook's worksheets into itself.</span></span> <span data-ttu-id="05a2b-135">O outro arquivo da pasta de trabalho é passado como em cadeia de caracteres codificado em base 64, como a chamada `Excel.createWorkbook`.</span><span class="sxs-lookup"><span data-stu-id="05a2b-135">The other workbook's file is passed as base64-encoded string, just like the `Excel.createWorkbook` call.</span></span>

```TypeScript
addFromBase64(base64File: string, sheetNamesToInsert?: string[], positionType?: Excel.WorksheetPositionType, relativeTo?: Worksheet | string): OfficeExtension.ClientResult<string[]>;
```

<span data-ttu-id="05a2b-136">O exemplo a seguir mostra planilhas da pasta de trabalho que estão sendo inseridas em uma pasta de trabalho atual, logo após a planilha ativa.</span><span class="sxs-lookup"><span data-stu-id="05a2b-136">The following example shows a workbook's worksheets being inserted in the current workbook, directly after the active worksheet.</span></span> <span data-ttu-id="05a2b-137">Observe que `null` é passado para o parâmetro `sheetNamesToInsert?: string[]`.</span><span class="sxs-lookup"><span data-stu-id="05a2b-137">Note that `null` is passed for the `sheetNamesToInsert?: string[]` parameter.</span></span> <span data-ttu-id="05a2b-138">Isso significa que todas as planilhas são inseridas.</span><span class="sxs-lookup"><span data-stu-id="05a2b-138">This means all the worksheets are being inserted.</span></span>

```js
var myFile = <HTMLInputElement>document.getElementById("file");
var reader = new FileReader();

reader.onload = (event) => {
    Excel.run((context) => {
        // strip off the metadata before the base64-encoded string
        var startIndex = (<string>(<FileReader>event.target).result).indexOf("base64,");
        var workbookContents = (<string>(<FileReader>event.target).result).substr(startIndex + 7);

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

## <a name="protect-the-workbooks-structure"></a><span data-ttu-id="05a2b-139">Protege a estrutura da pasta de trabalho</span><span class="sxs-lookup"><span data-stu-id="05a2b-139">Protect the workbook's structure</span></span>

<span data-ttu-id="05a2b-140">O suplemento pode controlar a capacidade de um usuário de editar dados em uma pasta de trabalho.</span><span class="sxs-lookup"><span data-stu-id="05a2b-140">Your add-in can control a user's ability to edit the workbook's structure.</span></span> <span data-ttu-id="05a2b-141">A propriedade `protection` do objeto Workbook é um objeto [WorkbookProtection](/javascript/api/excel/excel.workbookprotection) com um método `protect()`.</span><span class="sxs-lookup"><span data-stu-id="05a2b-141">The Workbook object's `protection` property is a [WorkbookProtection](/javascript/api/excel/excel.workbookprotection) object with a `protect()` method.</span></span> <span data-ttu-id="05a2b-142">O exemplo a seguir mostra um cenário básico ativando/desativando a proteção da pasta de trabalho.</span><span class="sxs-lookup"><span data-stu-id="05a2b-142">The following example shows a basic scenario toggling the protection of the workbook's structure.</span></span>

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

<span data-ttu-id="05a2b-143">O método `protect` aceita um parâmetro opcional de cadeia de caracteres.</span><span class="sxs-lookup"><span data-stu-id="05a2b-143">The `protect` method accepts an optional string parameter.</span></span> <span data-ttu-id="05a2b-144">Esta cadeia de caracteres representa a senha necessária para um usuário ignorar a proteção e alterar a estrutura da pasta de trabalho.</span><span class="sxs-lookup"><span data-stu-id="05a2b-144">This string represents the password needed for a user to bypass protection and change the workbook's structure.</span></span>

<span data-ttu-id="05a2b-145">A proteção também ser definida no nível da planilha para prevenir a edição de dados indesejada.</span><span class="sxs-lookup"><span data-stu-id="05a2b-145">Protection can also be set at the worksheet level to prevent unwanted data editing.</span></span> <span data-ttu-id="05a2b-146">Para saber mais, confira a seção **Proteção de dados**do artigo [Trabalhar com planilhas usando a API JavaScript do Excel](excel-add-ins-worksheets.md#data-protection).</span><span class="sxs-lookup"><span data-stu-id="05a2b-146">For more information, see the **Data protection** section of the [Work with worksheets using the Excel JavaScript API](excel-add-ins-worksheets.md#data-protection) article.</span></span>

> [!NOTE]
> <span data-ttu-id="05a2b-147">Para saber mais sobre a proteção de pastas de trabalho no Excel, confira o artigo [Proteger uma pasta de trabalho](https://support.office.com/article/Protect-a-workbook-7E365A4D-3E89-4616-84CA-1931257C1517).</span><span class="sxs-lookup"><span data-stu-id="05a2b-147">For more information about workbook protection in Excel, see the [Protect a workbook](https://support.office.com/article/Protect-a-workbook-7E365A4D-3E89-4616-84CA-1931257C1517) article.</span></span>

## <a name="access-document-properties"></a><span data-ttu-id="05a2b-148">Acessar propriedades do documentos</span><span class="sxs-lookup"><span data-stu-id="05a2b-148">Access document properties</span></span>

<span data-ttu-id="05a2b-149">Objetos Workbook têm acesso aos metadados dos arquivos do Office, que são conhecidos como [propriedades de documentos](https://support.office.com/article/View-or-change-the-properties-for-an-Office-file-21D604C2-481E-4379-8E54-1DD4622C6B75).</span><span class="sxs-lookup"><span data-stu-id="05a2b-149">Workbook objects have access to the Office file metadata, which is known as the [document properties](https://support.office.com/article/View-or-change-the-properties-for-an-Office-file-21D604C2-481E-4379-8E54-1DD4622C6B75).</span></span> <span data-ttu-id="05a2b-150">A propriedade `properties` do objeto Workbook é um objeto [DocumentProperties](/javascript/api/excel/excel.documentproperties) que contém esses valores de metadados.</span><span class="sxs-lookup"><span data-stu-id="05a2b-150">The Workbook object's `properties` property is a [DocumentProperties](/javascript/api/excel/excel.documentproperties) object containing these metadata values.</span></span> <span data-ttu-id="05a2b-151">O exemplo a seguir mostra como definir a propriedade **author**.</span><span class="sxs-lookup"><span data-stu-id="05a2b-151">The following example shows how to set the **author** property.</span></span>

```js
Excel.run(function (context) {
    var docProperties = context.workbook.properties;
    docProperties.author = "Alex";
    return context.sync();
}).catch(errorHandlerFunction);
```

<span data-ttu-id="05a2b-152">Você também pode definir propriedades personalizadas.</span><span class="sxs-lookup"><span data-stu-id="05a2b-152">You can also define custom properties.</span></span> <span data-ttu-id="05a2b-153">O objeto DocumentProperties contém uma propriedade `custom` que representa um conjunto de pares de valores-chave para propriedades definidas pelo usuário.</span><span class="sxs-lookup"><span data-stu-id="05a2b-153">The DocumentProperties object contains a `custom` property that represents a collection of key-value pairs for user-defined properties.</span></span> <span data-ttu-id="05a2b-154">O exemplo a seguir mostra como criar uma propriedade personalizada chamada **Introduction** com o valor "Olá" e, em seguida, recuperá-la.</span><span class="sxs-lookup"><span data-stu-id="05a2b-154">The following example shows how to create a custom property named **Introduction** with the value "Hello", then retrieve it.</span></span>

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

## <a name="access-document-settings"></a><span data-ttu-id="05a2b-155">Acessar configurações do documentos</span><span class="sxs-lookup"><span data-stu-id="05a2b-155">Access document settings</span></span>

<span data-ttu-id="05a2b-156">As configurações da pasta de trabalho são semelhantes ao conjunto de propriedades personalizadas.</span><span class="sxs-lookup"><span data-stu-id="05a2b-156">A workbook's settings are similar to the collection of custom properties.</span></span> <span data-ttu-id="05a2b-157">A diferença é que as configurações são exclusivas para um único arquivo do Excel e emparelhamento de suplementos, enquanto que as propriedades estão somente conectadas ao arquivo.</span><span class="sxs-lookup"><span data-stu-id="05a2b-157">The difference is settings are unique to a single Excel file and add-in pairing, whereas properties are solely connected to the file.</span></span> <span data-ttu-id="05a2b-158">O exemplo a seguir mostra como criar e acessar uma configuração.</span><span class="sxs-lookup"><span data-stu-id="05a2b-158">The following example shows how to create and access a setting.</span></span>

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

## <a name="add-custom-xml-data-to-the-workbook"></a><span data-ttu-id="05a2b-159">Adicionar dados XML personalizados à pasta de trabalho</span><span class="sxs-lookup"><span data-stu-id="05a2b-159">Add custom XML data to the workbook</span></span>

<span data-ttu-id="05a2b-160">O formato de arquivo Open XML **.xlsx** do Excel permite ao seu suplemento inserir dados XML personalizados na pasta de trabalho.</span><span class="sxs-lookup"><span data-stu-id="05a2b-160">Excel's Open XML **.xlsx** file format lets your add-in embed custom XML data in the workbook.</span></span> <span data-ttu-id="05a2b-161">Esses dados persistem na pasta de trabalho, independentemente do suplemento.</span><span class="sxs-lookup"><span data-stu-id="05a2b-161">This data persists with the workbook, independent of the add-in.</span></span>

<span data-ttu-id="05a2b-162">Uma pasta de trabalho contém um [CustomXmlPartCollection](/javascript/api/excel/excel.customxmlpartcollection), que é uma lista de [CustomXmlParts](/javascript/api/excel/excel.customxmlpart).</span><span class="sxs-lookup"><span data-stu-id="05a2b-162">A workbook contains a [CustomXmlPartCollection](/javascript/api/excel/excel.customxmlpartcollection), which is a list of [CustomXmlParts](/javascript/api/excel/excel.customxmlpart).</span></span> <span data-ttu-id="05a2b-163">Eles oferecem acesso a cadeias de caracteres XML e a uma ID exclusiva correspondente.</span><span class="sxs-lookup"><span data-stu-id="05a2b-163">These give access to the XML strings and a corresponding unique ID.</span></span> <span data-ttu-id="05a2b-164">Armazenando essas IDs como configurações, seu suplemento pode manter as teclas para suas partes XML entre sessões.</span><span class="sxs-lookup"><span data-stu-id="05a2b-164">By storing these IDs as settings, your add-in can maintain the keys to its XML parts between sessions.</span></span>

<span data-ttu-id="05a2b-165">Os exemplos a seguir mostram como usar partes XML personalizadas.</span><span class="sxs-lookup"><span data-stu-id="05a2b-165">The following samples show how to use custom XML parts.</span></span> <span data-ttu-id="05a2b-166">O primeiro bloco de códigos demonstra como inserir dados XML no documento.</span><span class="sxs-lookup"><span data-stu-id="05a2b-166">The first code block demonstrates how to embed XML data in the document.</span></span> <span data-ttu-id="05a2b-167">Ele armazena uma lista de revisores e usa as configurações da pasta de trabalho para salvar a `id` do XML para recuperação futura.</span><span class="sxs-lookup"><span data-stu-id="05a2b-167">It stores a list of reviewers, then uses the workbook's settings to save the XML's `id` for future retrieval.</span></span> <span data-ttu-id="05a2b-168">O segundo bloco mostra como acessar esse XML mais tarde.</span><span class="sxs-lookup"><span data-stu-id="05a2b-168">The second block shows how to access that XML later.</span></span> <span data-ttu-id="05a2b-169">A configuração "ContosoReviewXmlPartId" é carregada e transmitida para `customXmlParts` da pasta de trabalho.</span><span class="sxs-lookup"><span data-stu-id="05a2b-169">The "ContosoReviewXmlPartId" setting is loaded and passed to the workbook's `customXmlParts`.</span></span> <span data-ttu-id="05a2b-170">Os dados XML são então impressos no console.</span><span class="sxs-lookup"><span data-stu-id="05a2b-170">The XML data is then printed to the console.</span></span>

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
> <span data-ttu-id="05a2b-171">`CustomXMLPart.namespaceUri` só será preenchido se o elemento XML personalizado de nível superior contiver o atributo `xmlns`.</span><span class="sxs-lookup"><span data-stu-id="05a2b-171">`CustomXMLPart.namespaceUri` is only populated if the top-level custom XML element contains the `xmlns` attribute.</span></span>

## <a name="control-calculation-behavior"></a><span data-ttu-id="05a2b-172">Controlar o comportamento do cálculo</span><span class="sxs-lookup"><span data-stu-id="05a2b-172">Control calculation behavior</span></span>

### <a name="set-calculation-mode"></a><span data-ttu-id="05a2b-173">Configurar o modo de cálculo</span><span class="sxs-lookup"><span data-stu-id="05a2b-173">Set calculation mode</span></span>

<span data-ttu-id="05a2b-174">Por padrão, o Excel recalcula os resultados das fórmulas sempre que uma célula referenciada é alterada.</span><span class="sxs-lookup"><span data-stu-id="05a2b-174">By default, Excel recalculates formula results whenever a referenced cell is changed.</span></span> <span data-ttu-id="05a2b-175">O desempenho de seu suplemento pode se beneficiar do ajuste desse comportamento de cálculo.</span><span class="sxs-lookup"><span data-stu-id="05a2b-175">Your add-in's performance may benefit from adjusting this calculation behavior.</span></span> <span data-ttu-id="05a2b-176">O objeto Application tem uma propriedade `calculationMode` do tipo `CalculationMode`.</span><span class="sxs-lookup"><span data-stu-id="05a2b-176">The Application object has a `calculationMode` property of type `CalculationMode`.</span></span> <span data-ttu-id="05a2b-177">Esta propriedade pode ser configurada com os seguintes valores:</span><span class="sxs-lookup"><span data-stu-id="05a2b-177">It can be set to the following values:</span></span>

- <span data-ttu-id="05a2b-178">`automatic`: O comportamento de recálculo padrão em que o Excel calcula novos resultados das fórmulas sempre que o dado relevante é alterado.</span><span class="sxs-lookup"><span data-stu-id="05a2b-178">`automatic`: The default recalculation behavior where Excel calculates new formula results every time the relevant data is changed.</span></span>
- <span data-ttu-id="05a2b-179">`automaticExceptTables`: Igual a `automatic`, exceto que as alterações feitas nos valores em tabelas serão ignoradas.</span><span class="sxs-lookup"><span data-stu-id="05a2b-179">`automaticExceptTables`: Same as `automatic`, except any changes made to values in tables are ignored.</span></span>
- <span data-ttu-id="05a2b-180">`manual`: Os cálculos ocorrem somente quando o usuário ou suplemento os solicita.</span><span class="sxs-lookup"><span data-stu-id="05a2b-180">`manual`: Calculations only occur when the user or add-in requests them.</span></span>

### <a name="set-calculation-type"></a><span data-ttu-id="05a2b-181">Configurar o tipo de cálculo</span><span class="sxs-lookup"><span data-stu-id="05a2b-181">Set calculation type</span></span>

<span data-ttu-id="05a2b-182">O objeto [Application](/javascript/api/excel/excel.application) fornece um método para forçar um recálculo imediato.</span><span class="sxs-lookup"><span data-stu-id="05a2b-182">The [Application](/javascript/api/excel/excel.application) object provides a method to force an immediate recalculation.</span></span> <span data-ttu-id="05a2b-183">`Application.calculate(calculationType)` inicia o recálculo manual baseado no `calculationType` especificado.</span><span class="sxs-lookup"><span data-stu-id="05a2b-183">`Application.calculate(calculationType)` starts a manual recalculation based on the specified `calculationType`.</span></span> <span data-ttu-id="05a2b-184">Os seguintes valores podem ser especificados:</span><span class="sxs-lookup"><span data-stu-id="05a2b-184">The following values can be specified:</span></span>

- <span data-ttu-id="05a2b-185">`full`: Recalcule todas as fórmulas em todas as pastas de trabalho abertas, independentemente de elas terem sido alteradas desde o último recálculo.</span><span class="sxs-lookup"><span data-stu-id="05a2b-185">`full`: Recalculate all formulas in all open workbooks, regardless of whether they have changed since the last recalculation.</span></span>
- <span data-ttu-id="05a2b-186">`fullRebuild`: Verifique as fórmulas dependentes e depois recalcule todas as fórmulas em todas as pastas de trabalho abertas, independentemente de elas terem sido alteradas desde o último recálculo.</span><span class="sxs-lookup"><span data-stu-id="05a2b-186">`fullRebuild`: Check dependent formulas, and then recalculate all formulas in all open workbooks, regardless of whether they have changed since the last recalculation.</span></span>
- <span data-ttu-id="05a2b-187">`recalculate`: Recalcule as fórmulas que foram alteradas (ou marcadas por programação para recálculo) desde o último cálculo, e as fórmulas dependentes nelas, em todas as pastas de trabalho ativas.</span><span class="sxs-lookup"><span data-stu-id="05a2b-187">`recalculate`: Recalculate formulas that have changed (or been programmatically marked for recalculation) since the last calculation, and formulas dependent on them, in all active workbooks.</span></span>

> [!NOTE]
> <span data-ttu-id="05a2b-188">Para saber mais sobre o recálculo, confira o artigo [Alterar o recálculo, a iteração ou a precisão de fórmulas](https://support.office.com/article/change-formula-recalculation-iteration-or-precision-73fc7dac-91cf-4d36-86e8-67124f6bcce4).</span><span class="sxs-lookup"><span data-stu-id="05a2b-188">For more information about recalculation, see the [Change formula recalculation, iteration, or precision](https://support.office.com/article/change-formula-recalculation-iteration-or-precision-73fc7dac-91cf-4d36-86e8-67124f6bcce4) article.</span></span>

### <a name="temporarily-suspend-calculations"></a><span data-ttu-id="05a2b-189">Suspender os cálculos temporariamente</span><span class="sxs-lookup"><span data-stu-id="05a2b-189">Temporarily suspend calculations</span></span>

<span data-ttu-id="05a2b-190">A API do Excel também permite que os suplementos desativem os cálculos até que `RequestContext.sync()` seja chamado.</span><span class="sxs-lookup"><span data-stu-id="05a2b-190">The Excel API also lets add-ins turn off calculations until `RequestContext.sync()` is called.</span></span> <span data-ttu-id="05a2b-191">Isso é feito pelo `suspendApiCalculationUntilNextSync()`.</span><span class="sxs-lookup"><span data-stu-id="05a2b-191">This is done with `suspendApiCalculationUntilNextSync()`.</span></span> <span data-ttu-id="05a2b-192">Use esse método quando seu suplemento estiver editando intervalos extensos sem precisar acessar os dados entre as edições.</span><span class="sxs-lookup"><span data-stu-id="05a2b-192">Use this method when your add-in is editing large ranges without needing to access the data between edits.</span></span>

```js
context.application.suspendApiCalculationUntilNextSync();
```

## <a name="see-also"></a><span data-ttu-id="05a2b-193">Confira também</span><span class="sxs-lookup"><span data-stu-id="05a2b-193">See also</span></span>

- [<span data-ttu-id="05a2b-194">Conceitos fundamentais de programação com a API JavaScript do Excel</span><span class="sxs-lookup"><span data-stu-id="05a2b-194">Fundamental programming concepts with the Excel JavaScript API</span></span>](excel-add-ins-core-concepts.md)
- [<span data-ttu-id="05a2b-195">Trabalhar com planilhas usando a API JavaScript do Excel</span><span class="sxs-lookup"><span data-stu-id="05a2b-195">Work with worksheets using the Excel JavaScript API</span></span>](excel-add-ins-worksheets.md)
- [<span data-ttu-id="05a2b-196">Trabalhar com intervalos usando a API JavaScript do Excel</span><span class="sxs-lookup"><span data-stu-id="05a2b-196">Work with ranges using the Excel JavaScript API</span></span>](excel-add-ins-ranges.md)