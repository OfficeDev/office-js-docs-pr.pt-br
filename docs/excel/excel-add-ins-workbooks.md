---
title: Trabalhar com pastas de trabalho usando a API JavaScript do Excel
description: Exemplos de código que mostram como executar tarefas comuns com pastas de trabalho ou recursos de nível de aplicativo usando a API JavaScript do Excel.
ms.date: 03/19/2020
localization_priority: Normal
ms.openlocfilehash: aa30f888bf6de1926d2a36522febf0001e1e6130
ms.sourcegitcommit: 6c381634c77d316f34747131860db0a0bced2529
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 03/21/2020
ms.locfileid: "42891023"
---
# <a name="work-with-workbooks-using-the-excel-javascript-api"></a><span data-ttu-id="5e157-103">Trabalhar com pastas de trabalho usando a API JavaScript do Excel</span><span class="sxs-lookup"><span data-stu-id="5e157-103">Work with workbooks using the Excel JavaScript API</span></span>

<span data-ttu-id="5e157-104">Este artigo fornece exemplos de código que mostram como executar tarefas comuns com pastas de trabalho usando a API JavaScript do Excel.</span><span class="sxs-lookup"><span data-stu-id="5e157-104">This article provides code samples that show how to perform common tasks with workbooks using the Excel JavaScript API.</span></span> <span data-ttu-id="5e157-105">Para obter a lista completa de propriedades e métodos aos `Workbook` quais o objeto oferece suporte, consulte [objeto Workbook (API JavaScript para Excel)](/javascript/api/excel/excel.workbook).</span><span class="sxs-lookup"><span data-stu-id="5e157-105">For the complete list of properties and methods that the `Workbook` object supports, see [Workbook Object (JavaScript API for Excel)](/javascript/api/excel/excel.workbook).</span></span> <span data-ttu-id="5e157-106">Este artigo aborda também ações em nível de pasta de trabalho executadas através do objeto [Application](/javascript/api/excel/excel.application).</span><span class="sxs-lookup"><span data-stu-id="5e157-106">This article also covers workbook-level actions performed through the [Application](/javascript/api/excel/excel.application) object.</span></span>

<span data-ttu-id="5e157-107">O objeto Workbook é o ponto de entrada para que se suplemento interaja com o Excel.</span><span class="sxs-lookup"><span data-stu-id="5e157-107">The Workbook object is the entry point for your add-in to interact with Excel.</span></span> <span data-ttu-id="5e157-108">Ele mantém conjuntos de planilhas, tabelas, Tabelas Dinâmicas e muito mais, através dos quais os dados do Excel são acessados e alterados.</span><span class="sxs-lookup"><span data-stu-id="5e157-108">It maintains collections of worksheets, tables, PivotTables, and more, through which Excel data is accessed and changed.</span></span> <span data-ttu-id="5e157-109">O objeto [WorksheetCollection](/javascript/api/excel/excel.worksheetcollection) dá a seu suplemento acesso a todos os dados de pastas de trabalho através de planilhas individuais.</span><span class="sxs-lookup"><span data-stu-id="5e157-109">The [WorksheetCollection](/javascript/api/excel/excel.worksheetcollection) object gives your add-in access to all the workbook's data through individual worksheets.</span></span> <span data-ttu-id="5e157-110">Especificamente, ele permite seu suplemento adicione planilhas, navegue entre elas e atribua manipuladores a eventos de planilhas.</span><span class="sxs-lookup"><span data-stu-id="5e157-110">Specifically, it lets your add-in add worksheets, navigate among them, and assign handlers to worksheet events.</span></span> <span data-ttu-id="5e157-111">O artigo [Trabalhar com planilhas usando a API JavaScript do Excel](excel-add-ins-worksheets.md) descreve como acessar e editar planilhas.</span><span class="sxs-lookup"><span data-stu-id="5e157-111">The article [Work with worksheets using the Excel JavaScript API](excel-add-ins-worksheets.md) describes how to access and edit worksheets.</span></span>

## <a name="get-the-active-cell-or-selected-range"></a><span data-ttu-id="5e157-112">Obter a célula ativa ou o intervalo selecionado</span><span class="sxs-lookup"><span data-stu-id="5e157-112">Get the active cell or selected range</span></span>

<span data-ttu-id="5e157-113">O objeto Workbook contém dois métodos que obtêm um intervalo de células que o usuário ou o suplemento selecionaram: `getActiveCell()` e `getSelectedRange()`.</span><span class="sxs-lookup"><span data-stu-id="5e157-113">The Workbook object contains two methods that get a range of cells the user or add-in has selected: `getActiveCell()` and `getSelectedRange()`.</span></span> <span data-ttu-id="5e157-114">`getActiveCell()` obtém a célula ativa da pasta de trabalho como um [objeto Range](/javascript/api/excel/excel.range).</span><span class="sxs-lookup"><span data-stu-id="5e157-114">`getActiveCell()` gets the active cell from the workbook as a [Range object](/javascript/api/excel/excel.range).</span></span> <span data-ttu-id="5e157-115">O exemplo a seguir mostra uma chamada para `getActiveCell()`, seguida do endereço da célula que está sendo impresso no console.</span><span class="sxs-lookup"><span data-stu-id="5e157-115">The following example shows a call to `getActiveCell()`, followed by the cell's address being printed to the console.</span></span>

```js
Excel.run(function (context) {
    var activeCell = context.workbook.getActiveCell();
    activeCell.load("address");

    return context.sync().then(function () {
        console.log("The active cell is " + activeCell.address);
    });
}).catch(errorHandlerFunction);
```

<span data-ttu-id="5e157-116">O método `getSelectedRange()` retorna o intervalo único selecionado atualmente.</span><span class="sxs-lookup"><span data-stu-id="5e157-116">The `getSelectedRange()` method returns the currently selected single range.</span></span> <span data-ttu-id="5e157-117">Se houver vários intervalos selecionados, será gerado um erro InvalidSelection.</span><span class="sxs-lookup"><span data-stu-id="5e157-117">If multiple ranges are selected, an InvalidSelection error is thrown.</span></span> <span data-ttu-id="5e157-118">O exemplo a seguir mostra uma chamada para `getSelectedRange()` que, em seguida, define a cor de preenchimento do intervalo como amarelo.</span><span class="sxs-lookup"><span data-stu-id="5e157-118">The following example shows a call to `getSelectedRange()` that then sets the range's fill color to yellow.</span></span>

```js
Excel.run(function(context) {
    var range = context.workbook.getSelectedRange();
    range.format.fill.color = "yellow";
    return context.sync();
}).catch(errorHandlerFunction);
```

## <a name="create-a-workbook"></a><span data-ttu-id="5e157-119">Criar uma pasta de trabalho</span><span class="sxs-lookup"><span data-stu-id="5e157-119">Create a workbook</span></span>

<span data-ttu-id="5e157-120">O suplemento pode criar uma nova pasta de trabalho separada da instância do Excel, na qual o suplemento está sendo executado atualmente.</span><span class="sxs-lookup"><span data-stu-id="5e157-120">Your add-in can create a new workbook, separate from the Excel instance in which the add-in is currently running.</span></span> <span data-ttu-id="5e157-121">O objeto do Excel tem o método `createWorkbook` para esta finalidade.</span><span class="sxs-lookup"><span data-stu-id="5e157-121">The Excel object has the `createWorkbook` method for this purpose.</span></span> <span data-ttu-id="5e157-122">Quando esse método é chamado, a nova pasta de trabalho é aberta imediatamente e exibida em uma nova instância do Excel.</span><span class="sxs-lookup"><span data-stu-id="5e157-122">When this method is called, the new workbook is immediately opened and displayed in a new instance of Excel.</span></span> <span data-ttu-id="5e157-123">O suplemento permanece aberto e em execução com a pasta de trabalho anterior.</span><span class="sxs-lookup"><span data-stu-id="5e157-123">Your add-in remains open and running with the previous workbook.</span></span>

```js
Excel.createWorkbook();
```

<span data-ttu-id="5e157-124">O método `createWorkbook` também cria uma cópia de uma pasta de trabalho existente.</span><span class="sxs-lookup"><span data-stu-id="5e157-124">The `createWorkbook` method can also create a copy of an existing workbook.</span></span> <span data-ttu-id="5e157-125">O método aceita uma representação de cadeia de caracteres codificada em Base64 de um arquivo .xlsx como parâmetro opcional.</span><span class="sxs-lookup"><span data-stu-id="5e157-125">The method accepts a base64-encoded string representation of an .xlsx file as an optional parameter.</span></span> <span data-ttu-id="5e157-126">A pasta de trabalho resultante será uma cópia desse arquivo, supondo que o argumento da cadeia de caracteres seja um arquivo. xlsx válido.</span><span class="sxs-lookup"><span data-stu-id="5e157-126">The resulting workbook will be a copy of that file, assuming the string argument is a valid .xlsx file.</span></span>

<span data-ttu-id="5e157-127">Você pode obter a pasta de trabalho atual do suplemento como uma cadeia de caracteres codificada em base64 usando a [divisão de arquivos](/javascript/api/office/office.document#getfileasync-filetype--options--callback-).</span><span class="sxs-lookup"><span data-stu-id="5e157-127">You can get your add-in's current workbook as a base64-encoded string by using [file slicing](/javascript/api/office/office.document#getfileasync-filetype--options--callback-).</span></span> <span data-ttu-id="5e157-128">A classe [FileReader](https://developer.mozilla.org/docs/Web/API/FileReader) pode ser usada para converter um arquivo em uma cadeia de caracteres codificada com Base64, como demonstrado no seguinte exemplo.</span><span class="sxs-lookup"><span data-stu-id="5e157-128">The [FileReader](https://developer.mozilla.org/docs/Web/API/FileReader) class can be used to convert a file into the required base64-encoded string, as demonstrated in the following example.</span></span>

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

### <a name="insert-a-copy-of-an-existing-workbook-into-the-current-one-preview"></a><span data-ttu-id="5e157-129">Inserir uma cópia de uma pasta de trabalho para a seção atual (visualização)</span><span class="sxs-lookup"><span data-stu-id="5e157-129">Insert a copy of an existing workbook into the current one (preview)</span></span>

> [!NOTE]
> <span data-ttu-id="5e157-130">O método `WorksheetCollection.addFromBase64` no momento só está disponível na visualização pública e somente no Office no Windows e Mac.</span><span class="sxs-lookup"><span data-stu-id="5e157-130">The `WorksheetCollection.addFromBase64` method is currently only available in public preview and only for Office on Windows and Mac.</span></span> [!INCLUDE [Information about using preview APIs](../includes/using-excel-preview-apis.md)]

<span data-ttu-id="5e157-131">O exemplo anterior mostra uma nova pasta de trabalho criada a partir de uma pasta de trabalho.</span><span class="sxs-lookup"><span data-stu-id="5e157-131">The previous example shows a new workbook being created from an existing workbook.</span></span> <span data-ttu-id="5e157-132">Você também pode copiar algumas ou todas de uma pasta de trabalho para a atualmente associada com o suplemento.</span><span class="sxs-lookup"><span data-stu-id="5e157-132">You can also copy some or all of an existing workbook into the one currently associated with your add-in.</span></span> <span data-ttu-id="5e157-133">Uma pasta de trabalho [WorksheetCollection](/javascript/api/excel/excel.worksheetcollection) tem o método `addFromBase64` para inserir cópias de planilhas da pasta de trabalho de destino nela mesma.</span><span class="sxs-lookup"><span data-stu-id="5e157-133">A workbook's [WorksheetCollection](/javascript/api/excel/excel.worksheetcollection) has the `addFromBase64` method to insert copies of the target workbook's worksheets into itself.</span></span> <span data-ttu-id="5e157-134">O outro arquivo da pasta de trabalho é passado como em cadeia de caracteres codificado em base 64, como a chamada `Excel.createWorkbook`.</span><span class="sxs-lookup"><span data-stu-id="5e157-134">The other workbook's file is passed as base64-encoded string, just like the `Excel.createWorkbook` call.</span></span>

```TypeScript
addFromBase64(base64File: string, sheetNamesToInsert?: string[], positionType?: Excel.WorksheetPositionType, relativeTo?: Worksheet | string): OfficeExtension.ClientResult<string[]>;
```

<span data-ttu-id="5e157-135">O exemplo a seguir mostra planilhas da pasta de trabalho que estão sendo inseridas em uma pasta de trabalho atual, logo após a planilha ativa.</span><span class="sxs-lookup"><span data-stu-id="5e157-135">The following example shows a workbook's worksheets being inserted in the current workbook, directly after the active worksheet.</span></span> <span data-ttu-id="5e157-136">Observe que `null` é passado para o parâmetro `sheetNamesToInsert?: string[]`.</span><span class="sxs-lookup"><span data-stu-id="5e157-136">Note that `null` is passed for the `sheetNamesToInsert?: string[]` parameter.</span></span> <span data-ttu-id="5e157-137">Isso significa que todas as planilhas são inseridas.</span><span class="sxs-lookup"><span data-stu-id="5e157-137">This means all the worksheets are being inserted.</span></span>

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

## <a name="protect-the-workbooks-structure"></a><span data-ttu-id="5e157-138">Protege a estrutura da pasta de trabalho</span><span class="sxs-lookup"><span data-stu-id="5e157-138">Protect the workbook's structure</span></span>

<span data-ttu-id="5e157-139">O suplemento pode controlar a capacidade de um usuário de editar dados em uma pasta de trabalho.</span><span class="sxs-lookup"><span data-stu-id="5e157-139">Your add-in can control a user's ability to edit the workbook's structure.</span></span> <span data-ttu-id="5e157-140">A propriedade `protection` do objeto Workbook é um objeto [WorkbookProtection](/javascript/api/excel/excel.workbookprotection) com um método `protect()`.</span><span class="sxs-lookup"><span data-stu-id="5e157-140">The Workbook object's `protection` property is a [WorkbookProtection](/javascript/api/excel/excel.workbookprotection) object with a `protect()` method.</span></span> <span data-ttu-id="5e157-141">O exemplo a seguir mostra um cenário básico ativando/desativando a proteção da pasta de trabalho.</span><span class="sxs-lookup"><span data-stu-id="5e157-141">The following example shows a basic scenario toggling the protection of the workbook's structure.</span></span>

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

<span data-ttu-id="5e157-142">O método `protect` aceita um parâmetro opcional de cadeia de caracteres.</span><span class="sxs-lookup"><span data-stu-id="5e157-142">The `protect` method accepts an optional string parameter.</span></span> <span data-ttu-id="5e157-143">Esta cadeia de caracteres representa a senha necessária para um usuário ignorar a proteção e alterar a estrutura da pasta de trabalho.</span><span class="sxs-lookup"><span data-stu-id="5e157-143">This string represents the password needed for a user to bypass protection and change the workbook's structure.</span></span>

<span data-ttu-id="5e157-144">A proteção também ser definida no nível da planilha para prevenir a edição de dados indesejada.</span><span class="sxs-lookup"><span data-stu-id="5e157-144">Protection can also be set at the worksheet level to prevent unwanted data editing.</span></span> <span data-ttu-id="5e157-145">Para saber mais, confira a seção **Proteção de dados**do artigo [Trabalhar com planilhas usando a API JavaScript do Excel](excel-add-ins-worksheets.md#data-protection).</span><span class="sxs-lookup"><span data-stu-id="5e157-145">For more information, see the **Data protection** section of the [Work with worksheets using the Excel JavaScript API](excel-add-ins-worksheets.md#data-protection) article.</span></span>

> [!NOTE]
> <span data-ttu-id="5e157-146">Para saber mais sobre a proteção de pastas de trabalho no Excel, confira o artigo [Proteger uma pasta de trabalho](https://support.office.com/article/Protect-a-workbook-7E365A4D-3E89-4616-84CA-1931257C1517).</span><span class="sxs-lookup"><span data-stu-id="5e157-146">For more information about workbook protection in Excel, see the [Protect a workbook](https://support.office.com/article/Protect-a-workbook-7E365A4D-3E89-4616-84CA-1931257C1517) article.</span></span>

## <a name="access-document-properties"></a><span data-ttu-id="5e157-147">Acessar propriedades do documentos</span><span class="sxs-lookup"><span data-stu-id="5e157-147">Access document properties</span></span>

<span data-ttu-id="5e157-148">Objetos Workbook têm acesso aos metadados dos arquivos do Office, que são conhecidos como [propriedades de documentos](https://support.office.com/article/View-or-change-the-properties-for-an-Office-file-21D604C2-481E-4379-8E54-1DD4622C6B75).</span><span class="sxs-lookup"><span data-stu-id="5e157-148">Workbook objects have access to the Office file metadata, which is known as the [document properties](https://support.office.com/article/View-or-change-the-properties-for-an-Office-file-21D604C2-481E-4379-8E54-1DD4622C6B75).</span></span> <span data-ttu-id="5e157-149">A propriedade `properties` do objeto Workbook é um objeto [DocumentProperties](/javascript/api/excel/excel.documentproperties) que contém esses valores de metadados.</span><span class="sxs-lookup"><span data-stu-id="5e157-149">The Workbook object's `properties` property is a [DocumentProperties](/javascript/api/excel/excel.documentproperties) object containing these metadata values.</span></span> <span data-ttu-id="5e157-150">O exemplo a seguir mostra como definir a `author` propriedade.</span><span class="sxs-lookup"><span data-stu-id="5e157-150">The following example shows how to set the `author` property.</span></span>

```js
Excel.run(function (context) {
    var docProperties = context.workbook.properties;
    docProperties.author = "Alex";
    return context.sync();
}).catch(errorHandlerFunction);
```

<span data-ttu-id="5e157-151">Você também pode definir propriedades personalizadas.</span><span class="sxs-lookup"><span data-stu-id="5e157-151">You can also define custom properties.</span></span> <span data-ttu-id="5e157-152">O objeto DocumentProperties contém uma propriedade `custom` que representa um conjunto de pares de valores-chave para propriedades definidas pelo usuário.</span><span class="sxs-lookup"><span data-stu-id="5e157-152">The DocumentProperties object contains a `custom` property that represents a collection of key-value pairs for user-defined properties.</span></span> <span data-ttu-id="5e157-153">O exemplo a seguir mostra como criar uma propriedade personalizada chamada **Introduction** com o valor "Olá" e, em seguida, recuperá-la.</span><span class="sxs-lookup"><span data-stu-id="5e157-153">The following example shows how to create a custom property named **Introduction** with the value "Hello", then retrieve it.</span></span>

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

## <a name="access-document-settings"></a><span data-ttu-id="5e157-154">Acessar configurações do documentos</span><span class="sxs-lookup"><span data-stu-id="5e157-154">Access document settings</span></span>

<span data-ttu-id="5e157-155">As configurações da pasta de trabalho são semelhantes ao conjunto de propriedades personalizadas.</span><span class="sxs-lookup"><span data-stu-id="5e157-155">A workbook's settings are similar to the collection of custom properties.</span></span> <span data-ttu-id="5e157-156">A diferença é que as configurações são exclusivas para um único arquivo do Excel e emparelhamento de suplementos, enquanto que as propriedades estão somente conectadas ao arquivo.</span><span class="sxs-lookup"><span data-stu-id="5e157-156">The difference is settings are unique to a single Excel file and add-in pairing, whereas properties are solely connected to the file.</span></span> <span data-ttu-id="5e157-157">O exemplo a seguir mostra como criar e acessar uma configuração.</span><span class="sxs-lookup"><span data-stu-id="5e157-157">The following example shows how to create and access a setting.</span></span>

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

## <a name="access-application-culture-settings-preview"></a><span data-ttu-id="5e157-158">Configurações de cultura de aplicativo do Access (visualização)</span><span class="sxs-lookup"><span data-stu-id="5e157-158">Access application culture settings (preview)</span></span>

<span data-ttu-id="5e157-159">Uma pasta de trabalho tem configurações de idioma e cultura que afetam o modo como determinados dados são exibidos.</span><span class="sxs-lookup"><span data-stu-id="5e157-159">A workbook has language and culture settings that affect how certain data is displayed.</span></span> <span data-ttu-id="5e157-160">Essas configurações podem ajudar a localizar dados quando os usuários do seu suplemento estiverem compartilhando pastas de trabalho em diferentes idiomas e culturas.</span><span class="sxs-lookup"><span data-stu-id="5e157-160">These settings can help localize data when your add-in's users are sharing workbooks across different languages and cultures.</span></span> <span data-ttu-id="5e157-161">O suplemento pode usar a análise de cadeia de caracteres para localizar o formato de números, datas e horas com base nas configurações de cultura do sistema para que cada usuário veja os dados em seu próprio formato de cultura.</span><span class="sxs-lookup"><span data-stu-id="5e157-161">Your add-in can use string parsing to localize the format of numbers, dates, and times based on the system culture settings so that each user sees data in their own culture's format.</span></span>

<span data-ttu-id="5e157-162">`Application.cultureInfo`define as configurações de cultura do sistema como um objeto [CultureInfo](/javascript/api/excel/excel.cultureinfo) .</span><span class="sxs-lookup"><span data-stu-id="5e157-162">`Application.cultureInfo` defines the system culture settings as a [CultureInfo](/javascript/api/excel/excel.cultureinfo) object.</span></span> <span data-ttu-id="5e157-163">Contém configurações como o separador decimal numérico ou o formato de data.</span><span class="sxs-lookup"><span data-stu-id="5e157-163">This contains settings like the numerical decimal separator or the date format.</span></span>

<span data-ttu-id="5e157-164">Algumas configurações de cultura podem ser [alteradas por meio da interface do usuário do Excel](https://support.office.com/article/Change-the-character-used-to-separate-thousands-or-decimals-c093b545-71cb-4903-b205-aebb9837bd1e).</span><span class="sxs-lookup"><span data-stu-id="5e157-164">Some culture settings can be [changed through the Excel UI](https://support.office.com/article/Change-the-character-used-to-separate-thousands-or-decimals-c093b545-71cb-4903-b205-aebb9837bd1e).</span></span> <span data-ttu-id="5e157-165">As configurações do sistema são preservadas `CultureInfo` no objeto.</span><span class="sxs-lookup"><span data-stu-id="5e157-165">The system settings are preserved in the `CultureInfo` object.</span></span> <span data-ttu-id="5e157-166">As alterações locais são mantidas como propriedades no nível do [aplicativo](/javascript/api/excel/excel.application), `Application.decimalSeparator`como.</span><span class="sxs-lookup"><span data-stu-id="5e157-166">Any local changes are kept as [Application](/javascript/api/excel/excel.application)-level properties, such as `Application.decimalSeparator`.</span></span>

<span data-ttu-id="5e157-167">O exemplo a seguir altera o caractere separador decimal de uma cadeia de caracteres numérica de um ', ' para o caractere usado pelas configurações do sistema.</span><span class="sxs-lookup"><span data-stu-id="5e157-167">The following sample changes the decimal separator character of a numerical string from a ',' to the character used by the system settings.</span></span>

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

## <a name="add-custom-xml-data-to-the-workbook"></a><span data-ttu-id="5e157-168">Adicionar dados XML personalizados à pasta de trabalho</span><span class="sxs-lookup"><span data-stu-id="5e157-168">Add custom XML data to the workbook</span></span>

<span data-ttu-id="5e157-169">O formato de arquivo Open XML **.xlsx** do Excel permite ao seu suplemento inserir dados XML personalizados na pasta de trabalho.</span><span class="sxs-lookup"><span data-stu-id="5e157-169">Excel's Open XML **.xlsx** file format lets your add-in embed custom XML data in the workbook.</span></span> <span data-ttu-id="5e157-170">Esses dados persistem na pasta de trabalho, independentemente do suplemento.</span><span class="sxs-lookup"><span data-stu-id="5e157-170">This data persists with the workbook, independent of the add-in.</span></span>

<span data-ttu-id="5e157-171">Uma pasta de trabalho contém um [CustomXmlPartCollection](/javascript/api/excel/excel.customxmlpartcollection), que é uma lista de [CustomXmlParts](/javascript/api/excel/excel.customxmlpart).</span><span class="sxs-lookup"><span data-stu-id="5e157-171">A workbook contains a [CustomXmlPartCollection](/javascript/api/excel/excel.customxmlpartcollection), which is a list of [CustomXmlParts](/javascript/api/excel/excel.customxmlpart).</span></span> <span data-ttu-id="5e157-172">Eles oferecem acesso a cadeias de caracteres XML e a uma ID exclusiva correspondente.</span><span class="sxs-lookup"><span data-stu-id="5e157-172">These give access to the XML strings and a corresponding unique ID.</span></span> <span data-ttu-id="5e157-173">Armazenando essas IDs como configurações, seu suplemento pode manter as teclas para suas partes XML entre sessões.</span><span class="sxs-lookup"><span data-stu-id="5e157-173">By storing these IDs as settings, your add-in can maintain the keys to its XML parts between sessions.</span></span>

<span data-ttu-id="5e157-174">Os exemplos a seguir mostram como usar partes XML personalizadas.</span><span class="sxs-lookup"><span data-stu-id="5e157-174">The following samples show how to use custom XML parts.</span></span> <span data-ttu-id="5e157-175">O primeiro bloco de códigos demonstra como inserir dados XML no documento.</span><span class="sxs-lookup"><span data-stu-id="5e157-175">The first code block demonstrates how to embed XML data in the document.</span></span> <span data-ttu-id="5e157-176">Ele armazena uma lista de revisores e usa as configurações da pasta de trabalho para salvar a `id` do XML para recuperação futura.</span><span class="sxs-lookup"><span data-stu-id="5e157-176">It stores a list of reviewers, then uses the workbook's settings to save the XML's `id` for future retrieval.</span></span> <span data-ttu-id="5e157-177">O segundo bloco mostra como acessar esse XML mais tarde.</span><span class="sxs-lookup"><span data-stu-id="5e157-177">The second block shows how to access that XML later.</span></span> <span data-ttu-id="5e157-178">A configuração "ContosoReviewXmlPartId" é carregada e transmitida para `customXmlParts` da pasta de trabalho.</span><span class="sxs-lookup"><span data-stu-id="5e157-178">The "ContosoReviewXmlPartId" setting is loaded and passed to the workbook's `customXmlParts`.</span></span> <span data-ttu-id="5e157-179">Os dados XML são então impressos no console.</span><span class="sxs-lookup"><span data-stu-id="5e157-179">The XML data is then printed to the console.</span></span>

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
> <span data-ttu-id="5e157-180">`CustomXMLPart.namespaceUri` só será preenchido se o elemento XML personalizado de nível superior contiver o atributo `xmlns`.</span><span class="sxs-lookup"><span data-stu-id="5e157-180">`CustomXMLPart.namespaceUri` is only populated if the top-level custom XML element contains the `xmlns` attribute.</span></span>

## <a name="control-calculation-behavior"></a><span data-ttu-id="5e157-181">Controlar o comportamento do cálculo</span><span class="sxs-lookup"><span data-stu-id="5e157-181">Control calculation behavior</span></span>

### <a name="set-calculation-mode"></a><span data-ttu-id="5e157-182">Configurar o modo de cálculo</span><span class="sxs-lookup"><span data-stu-id="5e157-182">Set calculation mode</span></span>

<span data-ttu-id="5e157-183">Por padrão, o Excel recalcula os resultados das fórmulas sempre que uma célula referenciada é alterada.</span><span class="sxs-lookup"><span data-stu-id="5e157-183">By default, Excel recalculates formula results whenever a referenced cell is changed.</span></span> <span data-ttu-id="5e157-184">O desempenho de seu suplemento pode se beneficiar do ajuste desse comportamento de cálculo.</span><span class="sxs-lookup"><span data-stu-id="5e157-184">Your add-in's performance may benefit from adjusting this calculation behavior.</span></span> <span data-ttu-id="5e157-185">O objeto Application tem uma propriedade `calculationMode` do tipo `CalculationMode`.</span><span class="sxs-lookup"><span data-stu-id="5e157-185">The Application object has a `calculationMode` property of type `CalculationMode`.</span></span> <span data-ttu-id="5e157-186">Esta propriedade pode ser configurada com os seguintes valores:</span><span class="sxs-lookup"><span data-stu-id="5e157-186">It can be set to the following values:</span></span>

- <span data-ttu-id="5e157-187">`automatic`: O comportamento de recálculo padrão em que o Excel calcula novos resultados das fórmulas sempre que o dado relevante é alterado.</span><span class="sxs-lookup"><span data-stu-id="5e157-187">`automatic`: The default recalculation behavior where Excel calculates new formula results every time the relevant data is changed.</span></span>
- <span data-ttu-id="5e157-188">`automaticExceptTables`: Igual a `automatic`, exceto que as alterações feitas nos valores em tabelas serão ignoradas.</span><span class="sxs-lookup"><span data-stu-id="5e157-188">`automaticExceptTables`: Same as `automatic`, except any changes made to values in tables are ignored.</span></span>
- <span data-ttu-id="5e157-189">`manual`: Os cálculos ocorrem somente quando o usuário ou suplemento os solicita.</span><span class="sxs-lookup"><span data-stu-id="5e157-189">`manual`: Calculations only occur when the user or add-in requests them.</span></span>

### <a name="set-calculation-type"></a><span data-ttu-id="5e157-190">Configurar o tipo de cálculo</span><span class="sxs-lookup"><span data-stu-id="5e157-190">Set calculation type</span></span>

<span data-ttu-id="5e157-191">O objeto [Application](/javascript/api/excel/excel.application) fornece um método para forçar um recálculo imediato.</span><span class="sxs-lookup"><span data-stu-id="5e157-191">The [Application](/javascript/api/excel/excel.application) object provides a method to force an immediate recalculation.</span></span> <span data-ttu-id="5e157-192">`Application.calculate(calculationType)` inicia o recálculo manual baseado no `calculationType` especificado.</span><span class="sxs-lookup"><span data-stu-id="5e157-192">`Application.calculate(calculationType)` starts a manual recalculation based on the specified `calculationType`.</span></span> <span data-ttu-id="5e157-193">Os seguintes valores podem ser especificados:</span><span class="sxs-lookup"><span data-stu-id="5e157-193">The following values can be specified:</span></span>

- <span data-ttu-id="5e157-194">`full`: Recalcule todas as fórmulas em todas as pastas de trabalho abertas, independentemente de elas terem sido alteradas desde o último recálculo.</span><span class="sxs-lookup"><span data-stu-id="5e157-194">`full`: Recalculate all formulas in all open workbooks, regardless of whether they have changed since the last recalculation.</span></span>
- <span data-ttu-id="5e157-195">`fullRebuild`: Verifique as fórmulas dependentes e depois recalcule todas as fórmulas em todas as pastas de trabalho abertas, independentemente de elas terem sido alteradas desde o último recálculo.</span><span class="sxs-lookup"><span data-stu-id="5e157-195">`fullRebuild`: Check dependent formulas, and then recalculate all formulas in all open workbooks, regardless of whether they have changed since the last recalculation.</span></span>
- <span data-ttu-id="5e157-196">`recalculate`: Recalcule as fórmulas que foram alteradas (ou marcadas por programação para recálculo) desde o último cálculo, e as fórmulas dependentes nelas, em todas as pastas de trabalho ativas.</span><span class="sxs-lookup"><span data-stu-id="5e157-196">`recalculate`: Recalculate formulas that have changed (or been programmatically marked for recalculation) since the last calculation, and formulas dependent on them, in all active workbooks.</span></span>

> [!NOTE]
> <span data-ttu-id="5e157-197">Para saber mais sobre o recálculo, confira o artigo [Alterar o recálculo, a iteração ou a precisão de fórmulas](https://support.office.com/article/change-formula-recalculation-iteration-or-precision-73fc7dac-91cf-4d36-86e8-67124f6bcce4).</span><span class="sxs-lookup"><span data-stu-id="5e157-197">For more information about recalculation, see the [Change formula recalculation, iteration, or precision](https://support.office.com/article/change-formula-recalculation-iteration-or-precision-73fc7dac-91cf-4d36-86e8-67124f6bcce4) article.</span></span>

### <a name="temporarily-suspend-calculations"></a><span data-ttu-id="5e157-198">Suspender os cálculos temporariamente</span><span class="sxs-lookup"><span data-stu-id="5e157-198">Temporarily suspend calculations</span></span>

<span data-ttu-id="5e157-199">A API do Excel também permite que os suplementos desativem os cálculos até que `RequestContext.sync()` seja chamado.</span><span class="sxs-lookup"><span data-stu-id="5e157-199">The Excel API also lets add-ins turn off calculations until `RequestContext.sync()` is called.</span></span> <span data-ttu-id="5e157-200">Isso é feito pelo `suspendApiCalculationUntilNextSync()`.</span><span class="sxs-lookup"><span data-stu-id="5e157-200">This is done with `suspendApiCalculationUntilNextSync()`.</span></span> <span data-ttu-id="5e157-201">Use esse método quando seu suplemento estiver editando intervalos extensos sem precisar acessar os dados entre as edições.</span><span class="sxs-lookup"><span data-stu-id="5e157-201">Use this method when your add-in is editing large ranges without needing to access the data between edits.</span></span>

```js
context.application.suspendApiCalculationUntilNextSync();
```

## <a name="save-the-workbook-preview"></a><span data-ttu-id="5e157-202">Salvar a pasta de trabalho (visualização)</span><span class="sxs-lookup"><span data-stu-id="5e157-202">Save the workbook (preview)</span></span>

> [!NOTE]
> <span data-ttu-id="5e157-203">O método`Workbook.save` só está atualmente disponível na versão prévia pública.</span><span class="sxs-lookup"><span data-stu-id="5e157-203">The `Workbook.save` method is currently available only in public preview.</span></span> [!INCLUDE [Information about using preview APIs](../includes/using-excel-preview-apis.md)]

<span data-ttu-id="5e157-204">`Workbook.save` salva a pasta de trabalho para armazenamento persistente.</span><span class="sxs-lookup"><span data-stu-id="5e157-204">`Workbook.save` saves the workbook to persistent storage.</span></span> <span data-ttu-id="5e157-205">O método `save` usa um parâmetro simples e opcional `saveBehavior` que pode ter um dos seguintes valores:</span><span class="sxs-lookup"><span data-stu-id="5e157-205">The `save` method takes a single, optional `saveBehavior` parameter that can be one of the following values:</span></span>

- <span data-ttu-id="5e157-206">`Excel.SaveBehavior.save` (padrão): o arquivo será salvo sem solicitar que o usuário especifique o nome do arquivo e local de salvamento.</span><span class="sxs-lookup"><span data-stu-id="5e157-206">`Excel.SaveBehavior.save` (default): The file is saved without prompting the user to specify file name and save location.</span></span> <span data-ttu-id="5e157-207">Se o arquivo não tiver sido salvo anteriormente, ele será salvo no local padrão.</span><span class="sxs-lookup"><span data-stu-id="5e157-207">If the file has not been saved previously, it's saved to the default location.</span></span> <span data-ttu-id="5e157-208">Se o arquivo tiver sido salvo anteriormente, ele será salvo no mesmo local.</span><span class="sxs-lookup"><span data-stu-id="5e157-208">If the file has been saved previously, it's saved to the same location.</span></span>
- <span data-ttu-id="5e157-209">`Excel.SaveBehavior.prompt`: se o arquivo ainda não foi salvo anteriormente, o usuário será solicitado a especificar o nome do arquivo e o local de salvamento.</span><span class="sxs-lookup"><span data-stu-id="5e157-209">`Excel.SaveBehavior.prompt`: If file has not been saved previously, the user will be prompted to specify file name and save location.</span></span> <span data-ttu-id="5e157-210">Se o arquivo tiver sido salvo anteriormente, ele será salvo no mesmo local sem que o usuário seja solicitado.</span><span class="sxs-lookup"><span data-stu-id="5e157-210">If the file has been saved previously, it will be saved to the same location and the user will not be prompted.</span></span>

> [!CAUTION]
> <span data-ttu-id="5e157-211">Se o usuário for solicitado a salvar e, em vez disso, cancelar a operação, `save` gera uma exceção.</span><span class="sxs-lookup"><span data-stu-id="5e157-211">If the user is prompted to save and cancels the operation, `save` throws an exception.</span></span>

```js
context.workbook.save(Excel.SaveBehavior.prompt);
```

## <a name="close-the-workbook-preview"></a><span data-ttu-id="5e157-212">Feche a pasta de trabalho (visualização)</span><span class="sxs-lookup"><span data-stu-id="5e157-212">Close the workbook (preview)</span></span>

> [!NOTE]
> <span data-ttu-id="5e157-213">O método`Workbook.close` só está atualmente disponível na versão prévia pública.</span><span class="sxs-lookup"><span data-stu-id="5e157-213">The `Workbook.close` method is currently available only in public preview.</span></span> [!INCLUDE [Information about using preview APIs](../includes/using-excel-preview-apis.md)]

<span data-ttu-id="5e157-214">`Workbook.close` fecha a pasta de trabalho, além de suplementos que estão associados com a pasta de trabalho (o aplicativo Excel permanece aberto).</span><span class="sxs-lookup"><span data-stu-id="5e157-214">`Workbook.close` closes the workbook, along with add-ins that are associated with the workbook (the Excel application remains open).</span></span> <span data-ttu-id="5e157-215">O método `close` usa um parâmetro simples e opcional `closeBehavior` que pode ter um dos seguintes valores:</span><span class="sxs-lookup"><span data-stu-id="5e157-215">The `close` method takes a single, optional `closeBehavior` parameter that can be one of the following values:</span></span>

- <span data-ttu-id="5e157-216">`Excel.CloseBehavior.save` (padrão): o arquivo será salvo antes de fechar.</span><span class="sxs-lookup"><span data-stu-id="5e157-216">`Excel.CloseBehavior.save` (default): The file is saved before closing.</span></span> <span data-ttu-id="5e157-217">Se o arquivo não tiver sido salvo anteriormente, o usuário será solicitado a especificar o nome do arquivo e o local para salvá-lo.</span><span class="sxs-lookup"><span data-stu-id="5e157-217">If the file has not been saved previously, the user will be prompted to specify file name and save location.</span></span>
- <span data-ttu-id="5e157-218">`Excel.CloseBehavior.skipSave`: o arquivo é fechado imediatamente, sem ser salvo.</span><span class="sxs-lookup"><span data-stu-id="5e157-218">`Excel.CloseBehavior.skipSave`: The file is immediately closed, without saving.</span></span> <span data-ttu-id="5e157-219">Quaisquer alterações não salvas serão perdidas.</span><span class="sxs-lookup"><span data-stu-id="5e157-219">Any unsaved changes will be lost.</span></span>

```js
context.workbook.close(Excel.CloseBehavior.save);
```

## <a name="see-also"></a><span data-ttu-id="5e157-220">Confira também</span><span class="sxs-lookup"><span data-stu-id="5e157-220">See also</span></span>

- [<span data-ttu-id="5e157-221">Conceitos fundamentais de programação com a API JavaScript do Excel</span><span class="sxs-lookup"><span data-stu-id="5e157-221">Fundamental programming concepts with the Excel JavaScript API</span></span>](excel-add-ins-core-concepts.md)
- [<span data-ttu-id="5e157-222">Trabalhar com planilhas usando a API JavaScript do Excel</span><span class="sxs-lookup"><span data-stu-id="5e157-222">Work with worksheets using the Excel JavaScript API</span></span>](excel-add-ins-worksheets.md)
- [<span data-ttu-id="5e157-223">Trabalhar com intervalos usando a API JavaScript do Excel</span><span class="sxs-lookup"><span data-stu-id="5e157-223">Work with ranges using the Excel JavaScript API</span></span>](excel-add-ins-ranges.md)
